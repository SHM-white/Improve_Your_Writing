// See https://aka.ms/new-console-template for more information
using Improve_Your_Writing_Core;

Console.WriteLine("Start");

DocumentSettings documentSettings = new DocumentSettings() { 
    FontName = "SHM_white的字",
    FontSize = 24,
    InputXlsxPath = "D:\\Desktop\\test.xlsx",
    OutputDocxPath = "D:\\Desktop\\test.docx", 
    StartAfterLine = 0 
};
Improve_Your_Writing_Core.Core.Run(documentSettings);
