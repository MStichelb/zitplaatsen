# Zelf een .exe bouwen van Zitplaatsen

Met dit project kan je een standalone Windows-applicatie (.exe) maken van zitplaatsen.py.
De .exe kan je delen met anderen zodat zij het programma kunnen gebruiken zonder Python te installeren.

## Vereisten

Python 3.10+ geïnstalleerd op je systeem.

pip moet werken vanuit de terminal. Controleer dit met:

python --version
pip --version


Installeer PyInstaller (eenmalig):

pip install pyinstaller

## Stap 1. Repo downloaden of clonen

Download deze repository of clone hem naar je computer.
De belangrijkste bestanden/mappen die je nodig hebt zijn:

zitplaatsen.py
icons/       # bevat alle icoontjes die het programma gebruikt

## Stap 2. Map plopper toevoegen

De map plopper kan niet via GitHub gedeeld worden (om licentie-/bestandsredenen).
Je moet die zelf lokaal voorzien en in de root van het project zetten:

projectmap/
│
├── zitplaatsen.py
├── icons/
└── plopper/   # zelf lokaal toevoegen!


Zorg dat de mapstructuur correct is, anders kan de .exe het niet vinden.

## Stap 3. .exe bouwen met PyInstaller

Open een terminal (CMD of PowerShell) in de projectmap en voer uit:

pyinstaller --noconfirm --onefile --windowed ^
  --add-data "icons;icons" ^
  --add-data "poppler;poppler" ^
  --icon=icons\app.ico ^
  zitplaatsen.py

Wat gebeurt er?

PyInstaller maakt enkele nieuwe mappen aan:

build/
dist/
zitplaatsen.spec

Het uitvoerbaar bestand (zitplaatsen.exe) komt in de map dist/ terecht.

## Stap 4. Programma starten

Dubbelklik op zitplaatsen.exe in de dist map.
Dit is de standalone versie van je programma.

Je kan dit bestand nu ook naar anderen doorsturen (let op: de mappen icons en plopper moeten meegeleverd worden als externe bestanden).

## Extra tips

Wil je een nieuw icoontje gebruiken voor de .exe zelf? Voeg --icon=icons\plopper.ico toe aan het PyInstaller-commando:

pyinstaller --onefile --windowed --icon=icons\plopper.ico zitplaatsen.py


Heb je eerder al een .spec bestand aangemaakt? Verwijder dat dan voor je opnieuw bouwt:

del zitplaatsen.spec
