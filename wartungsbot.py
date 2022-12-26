# -*-coding: utf-8 -*-
import dataclasses
import datetime
import logging
import os
import re

import yaml
import configparser

from mwclient import Site


@dataclasses.dataclass
class Termin:
    """
    Klasse für einzelne Rollenspieltermine
    """
    kampagne: str = ''
    datum: datetime.date = ''
    tag: str = ''
    zeit: str = ''
    ort: str = ''
    status: str = ''
    link: str = ''
    spieler: list = dataclasses.field(default_factory=list)
    zusagen: list = dataclasses.field(default_factory=list)
    kommentare: str = ''


class Wartungsbot:
    def __init__(self, config: str):
        # Verzeichnis wechseln wegen lokaler Pfade
        self.param = None
        self.rpg_wiki = None
        self.protokoll = None
        self.param_ende = None
        self.param_start = None
        self.param_seite = None
        self.user = None
        self.password = None
        self.log_level = None
        self.localhost = None
        self.static_config = None
        os.chdir(os.path.dirname(os.path.abspath(__file__)))

        self.konfig_laden(config)
        self.wiki_login()
        self.parametrisierung_laden()

    def konfig_laden(self, config):
        # Konfigurationsdatei einlesen
        self.static_config = configparser.ConfigParser()
        self.static_config.read(config)

        self.localhost = self.static_config.getboolean('TECH', 'localhost')
        self.log_level = self.static_config['TECH']['log']
        log_level_info = {'logging.DEBUG': logging.DEBUG,
                          'logging.INFO': logging.INFO,
                          'logging.WARNING': logging.WARNING,
                          'logging.ERROR': logging.ERROR,
                          }
        self.log_level = log_level_info.get(self.log_level, logging.ERROR)
        self.user = self.static_config['ZUGANGSDATEN']['benutzer']
        self.password = self.static_config['ZUGANGSDATEN']['passwort']
        self.param_seite = self.static_config['PARAMETER']['seite']
        self.param_start = self.static_config['PARAMETER']['start']
        self.param_ende = self.static_config['PARAMETER']['ende']
        self.protokoll = self.static_config['PROTOKOLL']['dateiname']
        logging.basicConfig(filename='wartungsbot.log', filemode='a', format='%(asctime)s %(levelname)s: %(message)s',
                            datefmt='%d.%m.%y %H:%M:%S', level=self.log_level)

        logging.debug(
            f"Lokale Konfiguration:\n{self.localhost=},{self.user=},{self.password=},"
            f"{self.param_seite=},{self.param_start=},{self.param_ende=}")

    def wiki_login(self):
        # Falls Bot auf dem Pi läuft, Verbindung nach localhost und Verzicht auf SSL-Sicherung
        if self.localhost:
            self.rpg_wiki = Site('localhost', path='/mediawiki/', reqs={'verify': False})
        else:
            self.rpg_wiki = Site('www.rollenspiel-wiki.de', path='/mediawiki/')
        self.rpg_wiki.login(self.user, self.password)

    def parametrisierung_laden(self):
        # Parametrisierung aus Wiki-Seite auslesen
        self.param = self.rpg_wiki.pages[self.param_seite].text()
        pattern = r'(' + self.param_start + ')(.*?)(' + self.param_ende + ')'
        try:
            self.param = yaml.load(re.search(pattern, self.param, flags=re.S).group(2), Loader=yaml.SafeLoader)
        except yaml.YAMLError as e:
            logging.exception(e)
        logging.debug(f"Konfiguration gelesen:\n{self.param}")

    def termine_bereinigen(self):
        """
        Führt alle Aufgaben der Wartungsbot-Routine durch. Aktuell: Bereinigen abgelaufener Termine
        :return:
        """
        termine = self.termine_abfragen()

        # Liste der abgelaufenen Termine ermitteln
        heute = datetime.datetime.today().date()
        termine = [termin for termin in termine if (heute - termin.datum).days >= self.param['TageVergangen']]

        if not termine:
            logging.info('Keine Termine gefunden, die älter als ' + str(self.param['TageVergangen']) + ' Tage sind.')
            return

        for termin in termine:
            logging.info(f"Bereinige Termin vom {termin.datum.strftime('%d.%m.%Y')} für {termin.kampagne}.")
            self.termin_bereinigen(termin)

    def termine_abfragen(self) -> [Termin]:
        """
        Führt eine Abfrage im SemanticMediaWiki durch, um eine Liste der angesetzten Termine zu ermitteln.
        :return: Liste der Termine
        """
        query = r"""[[terminAnzeigen::wahr]]
                    |mainlabel=-
                    |?TerminDatum
                    |?TerminLink
                    |?TerminStatus
                    |?TerminZeit
                    |?TerminOrt
                    |?TerminLinkName
                    |?TerminTag
                    |format=template
                    |template=TerminTabelleSMW"""

        termine = []

        termin = Termin()
        for ergebnis in self.rpg_wiki.ask(query):
            for title, data in ergebnis.items():
                termin.kampagne = data['TerminLinkName'][0]
                termin.status = data['TerminStatus'][0] if data['TerminStatus'] else ''
                for fmt in ('%d.%m.%y', '%d.%m.%Y'):
                    try:
                        termin.datum = datetime.datetime.strptime(data['TerminDatum'][0], fmt).date() \
                            if data['TerminDatum'] else datetime.date(9999, 12, 31)
                    except ValueError:
                        pass
                termin.tag = data['TerminTag'][0] if data['TerminTag'] else ''
                if data['TerminLink']:
                    termin.link = data['TerminLink'][0]
                if termin:
                    termine.append(dataclasses.replace(termin))
        return termine

    def termin_bereinigen(self, termin: Termin):
        """
        Entfernte einen vergangenen Termin im Wiki.
        :param termin: Termin, der bereinigt werden soll
        :return:
        """
        # Bereinigte Termine in Terminprotokoll ablegen
        with open(self.protokoll, 'a+') as f:
            f.write(f"\n{termin.datum.strftime('%d.%m.%y')};{termin.kampagne};{termin.status}")

        seite = self.rpg_wiki.pages[termin.link].text()
        ergebnis = re.sub(r'(\|Status=)(.*?)(\|)', r'\1' + r'\n\3', seite, flags=re.S)
        ergebnis = re.sub(r'(\|Zusagen=)(.*?)(\|Status=)', r'\1' + r'\n\3', ergebnis, flags=re.S)
        ergebnis = re.sub(r'(\|Uhrzeit=)(.*?)(\|Spieler=)', r'\1' + r'\n\3', ergebnis, flags=re.S)
        ergebnis = re.sub(r'(\|Wochentag=)(.*?)(\|Kampagne=)', r'\1' + r'\n\3', ergebnis, flags=re.S)
        ergebnis = re.sub(r'(\|Datum=)(.*?)(\|Wochentag=)', r'\1' + r'\n\3', ergebnis, flags=re.S)
        self.rpg_wiki.pages[termin.link].edit(
            ergebnis,
            f"Wartungsbot: Vergangenen Termin vom {termin.datum.strftime('%d.%m.%y')} entfernt.",
            minor=False, bot=True)


def main():
    wb = Wartungsbot('wartungsbot.conf')

    if not wb.param['Aktiv']:
        logging.info('Bot nicht aktiv.')
        return

    if wb.param['VergangeneTermineBereinigen']:
        wb.termine_bereinigen()
    else:
        logging.info('Terminbereinigung nicht aktiviert.')


if __name__ == '__main__':
    main()
