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


def termine_abfragen(wiki: Site) -> [Termin]:
    """
    Führt eine Abfrage im SemanticMediaWiki durch, um eine Liste der angesetzten Termine zu ermitteln.
    :param wiki: Site-Objekt des Wikis
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
    for ergebnis in wiki.ask(query):
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


def termin_bereinigen(termin: Termin, wiki: Site):
    """
    Entfernte einen vergangenen Termin im Wiki.
    :param termin: Termin, der bereinigt werden soll
    :param wiki: Site-Objekt der Verbindung zum Wiki
    :return:
    """
    # Bereinigte Termine in Terminprotokoll ablegen
    with open(protokoll, 'a+') as f:
        f.write(f"\n{termin.datum.strftime('%d.%m.%y')};{termin.kampagne};{termin.status}")

    seite = wiki.pages[termin.link].text()
    ergebnis = re.sub(r'(\|Status=)(.*?)(\|)', r'\1' + r'\n\3', seite, flags=re.S)
    ergebnis = re.sub(r'(\|Zusagen=)(.*?)(\|Status=)', r'\1' + r'\n\3', ergebnis, flags=re.S)
    ergebnis = re.sub(r'(\|Uhrzeit=)(.*?)(\|Spieler=)', r'\1' + r'\n\3', ergebnis, flags=re.S)
    ergebnis = re.sub(r'(\|Wochentag=)(.*?)(\|Kampagne=)', r'\1' + r'\n\3', ergebnis, flags=re.S)
    ergebnis = re.sub(r'(\|Datum=)(.*?)(\|Wochentag=)', r'\1' + r'\n\3', ergebnis, flags=re.S)
    wiki.pages[termin.link].edit(ergebnis,
                                 f"Wartungsbot: Vergangenen Termin vom {termin.datum.strftime('%d.%m.%y')} entfernt.",
                                 minor=False, bot=True)


def bot_aktivieren(wiki: Site):
    """
    Führt alle Aufgaben der Wartungsbot-Routine durch. Aktuell: Bereinigen abgelaufener Termine
    :param wiki: Site-Objekt der Wiki-Verbindung
    :return:
    """
    if not param['VergangeneTermineBereinigen']:
        logging.info('Terminbereinigung nicht aktiviert.')
        return

    termine = termine_abfragen(wiki)

    # Liste der abgelaufenen Termine ermitteln
    heute = datetime.datetime.today().date()
    termine = [termin for termin in termine if (heute - termin.datum).days >= param['TageVergangen']]

    if not termine:
        logging.info('Keine Termine gefunden, die älter als ' + str(param['TageVergangen']) + ' Tage sind.')
        return

    for termin in termine:
        logging.info(f"Bereinige Termin vom {termin.datum.strftime('%d.%m.%Y')} für {termin.kampagne}.")
        termin_bereinigen(termin, wiki)


if __name__ == '__main__':
    # Verzeichnis wechseln wegen lokaler Pfade
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # Konfigurationsdatei einlesen
    static_config = configparser.ConfigParser()
    static_config.read('wartungsbot.conf')

    localhost = static_config.getboolean('TECH', 'localhost')
    log_level = static_config['TECH']['log']
    log_level_info = {'logging.DEBUG': logging.DEBUG,
                      'logging.INFO': logging.INFO,
                      'logging.WARNING': logging.WARNING,
                      'logging.ERROR': logging.ERROR,
                      }
    log_level = log_level_info.get(log_level, logging.ERROR)
    user = static_config['ZUGANGSDATEN']['benutzer']
    password = static_config['ZUGANGSDATEN']['passwort']
    param_seite = static_config['PARAMETER']['seite']
    param_start = static_config['PARAMETER']['start']
    param_ende = static_config['PARAMETER']['ende']
    protokoll = static_config['PROTOKOLL']['dateiname']

    # Logging-Funktionalität aktivieren
    logging.basicConfig(filename='wartungsbot.log', filemode='a', format='%(asctime)s %(levelname)s: %(message)s',
                        datefmt='%d.%m.%y %H:%M:%S', level=log_level)

    logging.debug(
        f"Lokale Konfiguration:\n{localhost=},{user=},{password=},{param_seite=},{param_start=},{param_ende=}")

    # Falls Bot auf dem Pi läuft, Verbindung nach localhost und Verzicht auf SSL-Sicherung
    if localhost:
        rpg_wiki = Site('localhost', path='/mediawiki/', reqs={'verify': False})
    else:
        rpg_wiki = Site('www.rollenspiel-wiki.de', path='/mediawiki/')
    rpg_wiki.login(user, password)

    # Parametrisierung aus Wiki-Seite auslesen
    param = rpg_wiki.pages[param_seite].text()
    pattern = r'(' + param_start + ')(.*?)(' + param_ende + ')'
    try:
        param = yaml.load(re.search(pattern, param, flags=re.S).group(2), Loader=yaml.SafeLoader)
    except yaml.YAMLError as e:
        logging.exception(e)
    logging.debug(f"Konfiguration gelesen:\n{param}")

    if param['Aktiv']:
        bot_aktivieren(rpg_wiki)
    else:
        logging.info('Bot nicht aktiv.')
