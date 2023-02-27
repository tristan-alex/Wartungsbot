# -*-coding: utf-8 -*-
import dataclasses
import datetime as dt
import json
import logging
import os
import re
import sys

import requests
import yaml
import configparser
from mwclient import Site
from tabulate import tabulate

WOCHENTAGE = ("Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag")


def terminplan_mail(abonnent: str, tabelle: [str]) -> str:
    """
    Mail, die als wöchentlicher Terminplan verschickt wird.
    :param abonnent: Wiki-Benutzername des Empfängers
    :param tabelle: Inhalt des Terminplans
    :return: Formatierte Mail
    """
    return f"""
Lieber {abonnent},\n\n
hier eine Übersicht der angesetzten Rollenspiel-Termine der nächsten Woche:

{tabulate(tabelle, headers=["Wochentag", "Datum", "Kampagne", "Status", "Eigene Aussage"], tablefmt="grid")}

Viele Grüße,
Dein Wartungsbot
    """


@dataclasses.dataclass
class Kampagne:
    """
    Klasse für Kampagnen
    """
    name: str = ''
    spieler: list = dataclasses.field(default_factory=list)


@dataclasses.dataclass
class Termin:
    """
    Klasse für einzelne Rollenspieltermine
    """
    kampagne: Kampagne = None
    datum: dt.date = None
    online: str = ''
    tag: str = ''
    zeit: str = ''
    ort: str = ''
    status: str = ''
    link: str = ''
    zusagen: list = dataclasses.field(default_factory=list)


@dataclasses.dataclass
class Status:
    """
    Klasse für den Zusagestatus der Spieler an einzelnen Terminen
    """
    datum: dt.date = None
    zusagen: list = dataclasses.field(default_factory=list)
    absagen: list = dataclasses.field(default_factory=list)
    unsicher: list = dataclasses.field(default_factory=list)


class Wartungsbot:
    """
    Klasse für den Wartungsbot
    """

    def __init__(self, config: str):
        """
        Initialisiert den Wartungsbot
        :param config: Dateiname der Konfigurationsdatei
        """
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
        self.termine_geladen = False
        self.termine = None
        self.kampagnen = None
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
        self.kampagnen_pfad = None

        self.konfig_laden(config)
        self.wiki_login()
        self.parametrisierung_laden()

    @staticmethod
    def namen_auslesen(text: str, spieler: bool = True) -> [str]:
        """
        Liest aus der Wiki-Seite eines Termins die Liste der Spieler oder die Liste der Zusagen aus.
        :param text: string-Darstellung der Wiki-Seite
        :param spieler: Wenn True, dann wird Liste der Spieler ausgelesen; wenn False, dann Liste der Zusagen
        :return: Liste von strings, welche die Namen der Spieler oder der Zusagen enthält
        """
        ergebnis = []
        if spieler:
            start, ende = '|Spieler=', '|EMailVerteiler='
        else:
            start, ende = '|Zusagen=', '|Status='
        for zeile in text[text.find(start) + len(start):text.rfind(ende)].splitlines():
            if ':' in zeile:
                ergebnis.append((zeile[zeile.find(':') + 1:zeile.rfind('|')]))
            else:
                ergebnis.append(zeile.replace('*', '').strip())

        mapping = {'Tris': 'Tristan'}

        for name in mapping.keys():
            if name in ergebnis:
                ergebnis.remove(name)
                ergebnis.append(mapping[name])

        return sorted([s for s in ergebnis if s])

    def konfig_laden(self, config: str):
        """
        Lädt die Konfiguration des Wartungsbots aus einer lokalen Konfigurationsdatei.
        :param config: Dateiname der Konfigurationsdatei
        :return:
        """
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
        self.kampagnen_pfad = self.static_config['TECH']['kampagnen_pfad']
        logging.basicConfig(filename='wartungsbot.log', filemode='a', format='%(asctime)s %(levelname)s: %(message)s',
                            datefmt='%d.%m.%y %H:%M:%S', level=self.log_level)

        logging.debug(
            f"Lokale Konfiguration:\n{self.localhost=},{self.user=},{self.password=},"
            f"{self.param_seite=},{self.param_start=},{self.param_ende=}")

    def wiki_login(self):
        """
        Führt ein Login im Wiki aus.
        :return:
        """
        if self.localhost:
            self.rpg_wiki = Site('localhost', path='/mediawiki/', reqs={'verify': False})
        else:
            self.rpg_wiki = Site('www.rollenspiel-wiki.de', path='/mediawiki/')
        self.rpg_wiki.login(self.user, self.password)

    def parametrisierung_laden(self):
        """
        Lädt die Parametrisierung des Wartungsbots aus dem Wiki.
        :return:
        """
        self.param = self.rpg_wiki.pages[self.param_seite].text()
        pattern = r'(' + self.param_start + ')(.*?)(' + self.param_ende + ')'
        try:
            self.param = yaml.load(re.search(pattern, self.param, flags=re.S).group(2), Loader=yaml.SafeLoader)
            self.param['Abonnenten'] = self.param['Abonnenten'].replace(' ', '').split(',')
        except yaml.YAMLError as e:
            logging.exception(e)
        logging.debug(f"Konfiguration gelesen:\n{self.param}")

    def termine_bereinigen(self):
        """
        Bereinigt abgelaufene Termine.
        :return:
        """
        self.termine_abfragen()

        heute = dt.datetime.today().date()

        for termin in self.termine:
            seite = self.rpg_wiki.pages[termin.link].text()

            if (heute - termin.datum).days >= self.param['TageVergangen']:
                logging.info(f"Bereinige Termin vom {termin.datum.strftime('%d.%m.%Y')} für {termin.kampagne.name}.")
                with open(self.protokoll, 'a+') as f:
                    f.write(f"\n{termin.datum.strftime('%d.%m.%y')};{termin.kampagne.name};{termin.status}")

                ergebnis = re.sub(r'(\|Status=)(.*?)(\|)', r'\1' + r'\n\3', seite, flags=re.S)
                ergebnis = re.sub(r'(\|Zusagen=)(.*?)(\|Status=)', r'\1' + r'\n\3', ergebnis, flags=re.S)
                ergebnis = re.sub(r'(\|Uhrzeit=)(.*?)(\|Spieler=)', r'\1' + r'\n\3', ergebnis, flags=re.S)
                ergebnis = re.sub(r'(\|Wochentag=)(.*?)(\|Kampagne=)', r'\1' + r'\n\3', ergebnis, flags=re.S)
                ergebnis = re.sub(r'(\|Datum=)(.*?)(\|Wochentag=)', r'\1' + r'\n\3', ergebnis, flags=re.S)
                msg = f"Wartungsbot: Vergangenen Termin vom {termin.datum.strftime('%d.%m.%y')} entfernt."
                self.rpg_wiki.pages[termin.link].edit(ergebnis, msg, minor=False, bot=True)

            datum = dt.datetime.strftime(termin.datum, '%d.%m.%Y') if termin.datum != dt.date(9999, 12, 31) else ''
            ergebnis = re.sub(r'(\|Datum=)(.*?)(\|Wochentag=)', r'\g<1>' + datum + r'\n\g<3>', seite, flags=re.S)

            if seite != ergebnis:
                logging.info(f"Passe Datumsformat von {termin.kampagne.name} auf {datum} an.")
                msg = f"Wartungsbot: Datumsformat vom {termin.kampagne.name}-Termin angepasst."
                self.rpg_wiki.pages[termin.link].edit(ergebnis, msg, minor=False, bot=True)

            wochentag = WOCHENTAGE[termin.datum.weekday()]
            ergebnis = re.sub(r'(\|Wochentag=)(.*?)(\|Kampagne=)', r'\g<1>' + wochentag + r'\n\g<3>', seite, flags=re.S)

            if seite != ergebnis:
                logging.info(f"Passe Wochentag von {termin.kampagne.name} auf {wochentag} an.")
                msg = f"Wartungsbot: Wochentag vom {termin.kampagne.name}-Termin angepasst."
                self.rpg_wiki.pages[termin.link].edit(ergebnis, msg, minor=False, bot=True)

    def termine_abfragen(self):
        """
        Führt eine Abfrage im SemanticMediaWiki durch, um eine Liste der angesetzten Termine zu ermitteln.
        :return:
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
                    |?Online
                    |format=template
                    |template=TerminTabelleSMW"""

        termine = []
        kampagnen = []

        for ergebnis in self.rpg_wiki.ask(query):
            for title, data in ergebnis.items():

                kampagne = Kampagne()
                termin = Termin()

                if data['TerminLink']:
                    termin.link = data['TerminLink'][0]
                    termin_seite = self.rpg_wiki.pages[termin.link].text()
                    kampagne.spieler = self.namen_auslesen(termin_seite)
                    termin.zusagen = self.namen_auslesen(termin_seite, spieler=False)

                kampagne.name = data['TerminLinkName'][0]
                termin.kampagne = kampagne
                termin.status = data['TerminStatus'][0] if data['TerminStatus'] else ''

                for fmt in ('%d.%m.%y', '%d.%m.%Y'):
                    try:
                        termin.datum = dt.datetime.strptime(data['TerminDatum'][0], fmt).date() \
                            if data['TerminDatum'] else dt.date(9999, 12, 31)
                    except ValueError:
                        pass

                termin.tag = data['TerminTag'][0] if data['TerminTag'] else ''
                termin.online = data['Online'][0] if data['Online'] else 'Nein'

                if termin:
                    termine.append(termin)
                if kampagne:
                    kampagnen.append(kampagne)

        self.termine = termine
        self.kampagnen = kampagnen
        self.termine_geladen = True

    def terminplan_mailen(self):
        """
        Schickt eine Mail an alle Abonnenten, in der die Rollenspieltermine enthalten sind, die bis einschließlich
        folgendem Sonntag angesetzt sind.
        :return:
        """
        if not self.termine_geladen:
            self.termine_abfragen()

        heute = dt.datetime.today().date()
        montag = heute + dt.timedelta(days=(-heute.weekday() % 7))
        sonntag = montag + dt.timedelta(days=6)

        abonnenten = self.param['Abonnenten']

        for abonnent in abonnenten:
            termine = [termin for termin in self.termine
                       if termin.datum <= sonntag and abonnent in termin.spieler]

            if not self.termine:
                msg = f"""Lieber {abonnent},\n\nin der nächsten Woche habe ich keine Rollenspiel-Termine gefunden,
                bei denen du mitspielen würdest.\n\nViele Grüße\nDein Wartungsbot"""
            else:
                termine = sorted(termine, key=lambda x: str(x.datum))
                tabelle = [[termin.tag, termin.datum.strftime('%d.%m.%y'), termin.kampagne.name, termin.status,
                            f"zugesagt" if abonnent in termin.zusagen else f"nicht zugesagt"] for termin in termine]
                msg = terminplan_mail(abonnent, tabelle)
            try:
                self.rpg_wiki.email(abonnent, msg, 'Wöchentliche Rollenspiel-Terminübersicht', cc=False)
                logging.info(f"Terminplan an {abonnent} verschickt.")
            except Exception as e:
                logging.error(f"Versand an {abonnent} fehlgeschlagen: {e}")

    def zusagestatus(self, von: dt.date, bis: dt.date) -> [Status]:
        """
        Ermittelt für alle Daten zwischen von und bis den Zusagestatus aller Spieler auf Basis von
        Jeanettes Terminplanungs-Skript.
        :param von: Startdatum
        :param bis: Enddatum
        :return: Liste der Terminstatus
        """
        von = '"from": {"dd": ' + str(von.day) + ', "mm": ' + str(von.month) + ', "yyyy": ' + str(von.year) + '}'
        bis = '"to": {"dd": ' + str(bis.day) + ', "mm": ' + str(bis.month) + ', "yyyy": ' + str(bis.year) + '}'

        data = {'functionname': 'getDateContentInRange', 'arguments': '{' + von + ', ' + bis + '}'}

        if self.localhost:
            url = 'https://localhost/mediawiki/extensions/terminplanung/terminplanungapi.php'
            verify = False
        else:
            url = 'https://www.rollenspiel-wiki.de/mediawiki/extensions/terminplanung/terminplanungapi.php'
            verify = True
        try:
            ergebnisse = requests.post(url, data=data, verify=verify).json()['result']
        except Exception as e:
            logging.error(f"Fehler bei Abfrage des Zusagestatus: {e}")
            return []
        ret = []
        for ergebnis in ergebnisse:
            status = Status()
            status.datum = dt.datetime.strptime(ergebnis['date'], '%d.%m.%Y').date()
            if not ergebnis['content']:
                ret.append(status)
                continue

            for typ in ergebnis['content']:
                personen = ergebnis['content'][typ]
                ergebnis['content'][typ] = [person['name'] for person in personen]

            status.zusagen = sorted(ergebnis['content']['accept'])
            status.absagen = sorted(ergebnis['content']['decline'])
            status.unsicher = sorted(ergebnis['content']['uncertain'])
            ret.append(status)
        return ret

    def terminideen(self, delta: int) -> [dict]:
        """
        Berechnet Terminvorschläge auf Basis des Zusagestatus.
        :param delta: Anzahl der Tage, die in die Zukunft geschaut werden soll
        :return: Dictionary mit Kampagne, Datum und einer Liste der fehlenden Zusagen
        """
        heute = dt.datetime.today().date()
        enddatum = heute + dt.timedelta(days=delta)
        statusliste = self.zusagestatus(heute, enddatum)
        ret = []

        if not self.termine_geladen:
            self.termine_abfragen()
            self.termine_geladen = True
        verboten = [termin.datum for termin in self.termine if termin.status in ['Angesetzt', 'Bestätigt']]

        for kampagne in [kampagne for kampagne in self.kampagnen if kampagne.name != 'Testkampagne']:
            online = [termin.online for termin in self.termine if termin.kampagne == kampagne][0]
            datum = heute
            spieldatum = None
            while datum <= enddatum:

                if datum not in [status.datum for status in statusliste]:
                    break

                status = [status for status in statusliste if status.datum == datum][0]
                if set(status.absagen).intersection(set(kampagne.spieler)):
                    datum = datum + dt.timedelta(days=1)
                    continue

                if datum in verboten or (datum.weekday() <= 3 and online == 'Nein'):
                    datum = datum + dt.timedelta(days=1)
                    continue

                if set(kampagne.spieler).issubset(set(status.zusagen)):
                    spieldatum = datum
                    break

                if not spieldatum:
                    spieldatum = datum

                datum = datum + dt.timedelta(days=1)

            if spieldatum:
                status = [status for status in statusliste if status.datum == spieldatum][0]
                fehlende_zusagen = [spieler for spieler in kampagne.spieler if spieler not in status.zusagen]
            else:
                fehlende_zusagen = []

            ret.append({'Kampagne': kampagne.name, 'Datum': spieldatum, 'Fehlen': fehlende_zusagen})
        return ret

    def tabelle_formatieren(self, tabelle: [str]) -> str:
        """
        Formatiert die Tabelle für die Terminideen analog zur Terminübersicht auf der Hauptseite.
        :param tabelle: Inhalt der Tabelle.
        :return: Als String formatierte Tabelle im Mediawiki-Format analog zur Terminübersicht
        """
        headers = ['Kampagne', 'Tag', 'Datum', 'Fehlende Zusagen', 'Terminvorschlag?']
        sizes = ['12em', '8em', '10em', '25em', '12em']
        zuordnung = dict(zip(headers, sizes))
        farbe = {'Termin möglich!': '#CCFF66',
                 'Termin eventuell möglich.': '#FFFF66',
                 'kein Termin möglich.': '#FA5858'}
        erg = '<div style="float:left; float:left; background:none; font-size: 100%;">\n' \
              '{| cellspacing="0" cellpadding="5" style="border-collapse:collapse" class="sortable"'
        for header in headers:
            erg = erg + '\n!align="left" style="border-bottom: 2pt darkred solid; width:' \
                  + zuordnung[header] + '"|' + header
        for zeile in tabelle:
            erg = erg + '\n|-'
            for i, spalte in enumerate(zeile):
                stil = '\n|style="spacing-bottom: 0px; vertical-align: middle; border-bottom: 1pt lightgray solid;'
                if headers[i] == 'Terminvorschlag?':
                    stil = stil + 'color:#000000; background-color:' + farbe[zeile[i]]
                if headers[i] == 'Kampagne':
                    link = [termin.link for termin in self.termine if termin.kampagne.name == spalte][0]
                    spalte = '[[' + link + '|' + spalte + ']]'
                erg = erg + stil + '"|' + spalte
        erg = erg + '\n|-\n|}\n</div>'
        return erg

    def terminideen_posten(self, delta: int = 90):
        """
        Postet Terminideen im Tabellenformat auf der Hauptseite.
        :param delta: Anzahl der Tage, die in die Zukunft geschaut werden soll
        """
        ret = sorted(self.terminideen(delta), key=lambda x: x['Datum'])
        for kampagne in ret:
            if kampagne['Datum'] and not kampagne['Fehlen']:
                kampagne['Vorschlag'] = 'Termin möglich!'
            elif kampagne['Datum']:
                kampagne['Vorschlag'] = 'Termin eventuell möglich.'
            else:
                kampagne['Vorschlag'] = 'kein Termin möglich.'
        tabelle = [[kampagne['Kampagne'],
                    WOCHENTAGE[kampagne['Datum'].weekday()] if kampagne['Datum'] else '',
                    dt.datetime.strftime(kampagne['Datum'], '%d.%m.%Y') if kampagne['Datum'] else 'Keinen Termin '
                                                                                                  'gefunden',
                    ', '.join(kampagne['Fehlen']), kampagne['Vorschlag']] for kampagne in ret]
        wiki = self.tabelle_formatieren(tabelle)
        seite = self.rpg_wiki.pages['Hauptseite'].text()
        ergebnis = re.sub(r'(===Terminideen===)(.*?)(</div>)', r'\1\n' + wiki, seite, flags=re.S)
        if ergebnis == seite:
            logging.info(f"Keine Aktualisierung der Terminvorschläge erforderlich.")
        else:
            msg = f"Wartungsbot: Tabelle Terminvorschläge aktualisiert."
            self.rpg_wiki.pages['Hauptseite'].edit(ergebnis, msg, minor=False, bot=True)
            logging.info(f"Terminvorschläge für {delta} Vorschautage gepostet.")

    def kampagnen_synchronisieren(self, ausschluss: [str], pfad: str):
        """
        Legt eine Liste der Kampagnen mit Spielern im passenden Format für Jeanettes Terminplanungs-Tool ab.
        :param ausschluss: Liste von Kampagnennamen, die nicht mit abgelegt werden sollen (für Testzwecke)
        :param pfad: Pfad der Zieldatei
        :return:
        """
        if not self.termine_geladen:
            self.termine_abfragen()

        with open(pfad, 'w') as f:
            kampagnen = [{'name': termin.kampagne.name, 'player': termin.kampagne.spieler}
                         for termin in self.termine if termin.kampagne.name not in ausschluss]
            kampagnen = sorted(kampagnen, key=lambda x: x['name'])
            logging.info(f"Synchronisiere Kampagnen mit Terminplanung.")
            json.dump(kampagnen, f)


def main():
    """
    Einstiegspunkt des Skripts.
    :return:
    """
    if len(sys.argv) > 1:
        argument = sys.argv[1]
        erlaubte_argumente = ['terminplan']
        if argument not in erlaubte_argumente:
            logging.error(f'Argument nicht erkannt: {argument}. Erlaubt sind: {erlaubte_argumente}.')
            return
    else:
        argument = None

    wb = Wartungsbot('wartungsbot.conf')

    if not wb.param['Aktiv']:
        logging.info('Bot nicht aktiv.')
        return

    if not argument:
        if wb.param['VergangeneTermineBereinigen']:
            wb.termine_bereinigen()
        else:
            logging.info('Terminbereinigung nicht aktiviert.')

        if wb.param['TerminideenPosten']:
            wb.terminideen_posten(wb.param['TerminideenZeitfenster'])
        else:
            logging.info('Posten von Terminideen nicht aktiviert.')

        wb.kampagnen_synchronisieren(ausschluss=['Testkampagne'], pfad=wb.kampagnen_pfad + 'kampagnen.txt')
    else:
        logging.info("Standardfunktionen nicht aktiviert.")

    if wb.param['TerminplanVersenden'] and argument == 'terminplan':
        wb.terminplan_mailen()
    else:
        logging.info('Versand Terminplan nicht aktiviert.')


if __name__ == '__main__':
    main()
