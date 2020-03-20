# OpenXmlPlayground

Questa soluzione contiene:
- una DLL con al suo interno due classi di metodi statici per la semplificazione da parte del programmatore per la creazione di file MS Word e Excel.
- una applicazione Windows Form che prova ogni funzione con la creazione di un file .docx e .xslx di prova.

## Uso
Per utilizzare la DLL ci sarà bisogno dell'aggiunta del seguente codice:
```c#
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;
```

## Word.cs
Permette di creare un testo formattato e con stili personalizzati, inserire immagini nel documento, creare tabelle e elenchi puntati tramite metodi statici tra cui:
```c#
public static void InsertPicture(WordprocessingDocument wordprocessingDocument, string fileName)
public static Paragraph CreateParagraphWithStyle(string styleId, JustificationValues justification)
public static void CreateBulletNumberingPart(MainDocumentPart mainPart, string bulletChar = "-")
public static Table CreateTable(string[][] contenuto, TableProperties tableProperties)
```
Questa classe è più macchinosa e per creare un documento bisogna leggere bene i parametri dei metodi e ci potrebbe essere l'eventualità di dover ritoccare alcuni metodi per adattarla alla soluzione

## Excel.cs
Permette la facile creazione di uno spreadsheet tramite un solo metodo:
```c#
public static void CreateExcelFile<T>(List<T> data,string path)
```
Quest'ultimo necessita soltanto di una lista di dati e del path di destinazione del file e verrà creata una tabella che conterrà un'intestazione con i nomi di tutte le proprietà dell'oggetto e le sue proprietà in una riga della tabella.


Per domande inviate una mail a Alberto Bongioanni: a.bongioanni.0746@vallauri.edu
