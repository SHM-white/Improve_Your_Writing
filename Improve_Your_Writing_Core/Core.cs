using System.ComponentModel;
using System.Runtime.Versioning;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;


namespace Improve_Your_Writing_Core
{
    static public class Core
    {
        public static bool Run(DocumentSettings settings)
        {
            if (settings == null)
            {
                return false;
            }
            Dictionary<string, string> data = GetXlsxData(settings.InputXlsxPath);

            return WriteToDocx(settings.OutputDocxPath, data, settings);
        }

        private static bool WriteToDocx(string path, Dictionary<string, string> data, DocumentSettings settings)
        {
            // 创建一个新的DOCX文档对象
            XWPFDocument document = new XWPFDocument();

            // 添加段落
            XWPFParagraph paragraph = document.CreateParagraph();
            XWPFRun run = paragraph.CreateRun();
            run.SetText("Hello, World!");

            // 保存文档
            using (FileStream file = new FileStream(settings.OutputDocxPath, FileMode.Create, FileAccess.Write))
            {
                document.Write(file);
            }

            return true;
        }

        private static Dictionary<string, string> GetXlsxData(string path)
        {
            Dictionary<string, string> data = new Dictionary<string, string>();
            // 创建一个工作簿对象
            IWorkbook workbook;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);
            }

            // 获取第一个工作表
            ISheet sheet = workbook.GetSheetAt(0);

            // 遍历行
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    string[] strings = new string[row.LastCellNum];

                    // 遍历单元格
                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        NPOI.SS.UserModel.ICell cell = row.GetCell(j);
                        if (cell != null)
                        {
                            // 获取单元格的值
                            string value = cell.ToString();
                            // 处理单元格的值
                            strings[j] = value;

                        }
                    }
                    data.Add(strings[0], strings[1]);
                }
            }
            return data;
        }
    }


    public class DocumentSettings : INotifyPropertyChanged
    {

        public DocumentSettings()
        {

        }
        private int _fontSize;
        public int FontSize
        {
            get { return _fontSize; }
            set
            {
                if (_fontSize != value)
                {
                    _fontSize = value;
                    OnPropertyChanged(nameof(FontSize));
                }
            }
        }

        private string? _fontName;
        public string? FontName
        {
            get { return _fontName; }
            set
            {
                if (_fontName != value)
                {
                    _fontName = value;
                    OnPropertyChanged(nameof(FontName));
                }
            }
        }

        private string? _outputDocxPath;
        public string? OutputDocxPath
        {
            get { return _outputDocxPath; }
            set
            {
                if (_outputDocxPath != value)
                {
                    _outputDocxPath = value;
                    OnPropertyChanged(nameof(OutputDocxPath));
                }
            }
        }

        private string? _inputXlsxPath;
        public string? InputXlsxPath
        {
            get { return _inputXlsxPath; }
            set
            {
                if (_inputXlsxPath != value)
                {
                    _inputXlsxPath = value;
                    OnPropertyChanged(nameof(InputXlsxPath));
                }
            }
        }

        private int _startAfterLine;
        public int StartAfterLine
        {
            get { return _startAfterLine; }
            set
            {
                if (_startAfterLine != value)
                {
                    _startAfterLine = value;
                    OnPropertyChanged(nameof(StartAfterLine));
                }
            }
        }

        // INotifyPropertyChanged implementation
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }


    }
}