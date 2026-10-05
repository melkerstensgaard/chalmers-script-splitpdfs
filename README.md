### Guide/info för inskanningsarbetet
Detta dokument beskriver ett Python skript som tillsammans OCRmyPDF effektiviserar flera delar av inskanningsarbetet.
Bulletpoints över vad skriptet gör:
- Delar automatiskt intygen så att varje intyg blir en egen PDF
- Lägger in filnamn, volymnamn och personnummer i ett exceldokument.
- Extraherar personnummer ur innehållet i PDF:en, detta fungerar med alla olika format vi har hittat samt T nummer
o Detta gör den genom att kolla alla personnummer liknande sifferkombinationer som förekommer i varje bevis och väljer den som förekommer flest gånger, därmed blir det rätt även om OCR-skanningen inte träffade rätt på alla
o Om den inte hittar en 10-siffrig sifferkombination letar den efter andra format vi hittat, tex yyyy-mm-dd, yyyymmdd, yyyy-mmdd med mera.
- Rapporterar fel till en log fil
- Skippar hela volymen om någonting är fel med någon PDF däri, rapporterar sedan detta till en separat log fil

## Skärmdump
<img width="602" height="858" alt="image" src="https://github.com/user-attachments/assets/30522767-87ac-4214-bfac-51228dd1c505" />


### Användning
Innan man kan använda skriptet måste man ladda ner de pythonpaket som behövs. Dessa finns i requirements.txt. Installera dem genom följande kommando:

```python.exe -m pip install -r requirements.txt```

Först kan man dock behöva installera pip genom att köra

```python.exe -m ensurepip --upgrade```

För att göra OCR mha OCRmyPDF skriver man:

```exec ocrmypdf -l swe --output-type pdf <pdf>```

för att köra igenom hela mappar gör man:

```find . -name '*.pdf' -printf '%p\n' -exec ocrmypdf -l swe --output-type pdf '{}' '{}' \;```

För att använda bearbetningsskriptet gör man i nuläget följande i WSL:

```Python3 bearbetningsskript.py input/ output/```

Eller via gränssnittet via:

```Python3 gränssnitt.py ```

Gränssnittet aktiverar bara skriptet åt en men man kan också lägga till fler markörer däri så det finns lite mer funktionalitet om man använder gränssnittet.

### Workflow
1. Ta ut alla bevisen ur fasciklarna, låt bilagor ligga kvar.
2. Skanna filerna till USB så att de ligger i mappar med samma namn som volymen.
3. För över mapparna från USB till "Obearbetat".
4. Kör OCRmyPDF kommandot.
5. Kör python skriptet.
6. Spara till "bearbetat"
8. Validera.
9. Flytta till "färdigt".
Klart! 

### Beroenden 

-	[WSL]([url](https://learn.microsoft.com/en-us/windows/wsl/install)) (Windows Subsystem for Linux) – tillåter en begränsad Linuxmiljö i Windows som vi behöver för att använda OCRmyPDF. 
-	[OCRmyPDF]([url](https://github.com/ocrmypdf/OCRmyPDF)) – OCRverktyg, installeras i WSL. 
-	[Python 3]([url](https://www.python.org/downloads/windows/)) – använd den senaste 64-bit versionen. 
Andra verktyg vi använt är:
-	[VeraPDF]([url](https://docs.verapdf.org/install/)) (för PDF/A-2u granskning och verifiering, det är Greenfield vi använder)
-	[Java]([url](https://www.java.com/en/download/)) – VeraPDF behöver java
-	[PDFtoPDFa]([url](https://github.com/iRedPaul/pdftopdfa/blob/main/docs/usage.md)) – konvertera pdf till pdf/a2u


