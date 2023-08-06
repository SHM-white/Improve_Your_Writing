// See https://aka.ms/new-console-template for more information
using Improve_Your_Writing_Core;

Console.WriteLine("Strat");

DocumentSettings documentSettings = new DocumentSettings() { 
    FontName = "SHM_white的字",
    FontSize = 20,
    InputXlsxPath = "test.xlsx",
    OutputDocxPath = "test.docx", 
    StartAfterLine = 0 
};
Improve_Your_Writing_Core.Core.Run(documentSettings);
