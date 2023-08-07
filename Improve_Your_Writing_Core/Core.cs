using System.ComponentModel;
using System.Runtime.Versioning;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using System.Collections.ObjectModel;
using System.Drawing.Text;

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
            paragraph.SpacingLineRule = LineSpacingRule.EXACT;
            paragraph.SpacingBetween = 5;

            XWPFRun run1 = paragraph.CreateRun();
            run1.SetText("\n\n\n\n\n\n\n\n\n");
            run1.FontSize = settings.FontSize;
            run1.FontFamily = "等线";
            run1.AddBreak(BreakType.TEXTWRAPPING);

            XWPFRun run = paragraph.CreateRun();
            run.SetText("Aa Bb Cc Dd Ee Ff Gg Hh Ii Jj Kk Ll Mm Nn Oo Pp Qq Rr Ss Tt Uu Vv Ww Xx Yy Zz");
            run.FontFamily = settings.FontName;
            run.FontSize = settings.FontSize;
            run.AddBreak(BreakType.TEXTWRAPPING);

            XWPFParagraph paragraph2 = document.CreateParagraph();
            paragraph2.SpacingLineRule = LineSpacingRule.EXACT;
            paragraph2.SpacingBetween = 28;
            foreach(var item in data)
            {
                XWPFRun run2 = paragraph2.CreateRun();
                run2.SetText(item.Key+"  "+item.Key+"  "+item.Key+"         ");
                run2.FontFamily = settings.FontName;
                run2.FontSize = settings.FontSize;
            }

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
            _fonts = new();
            GetFonts();
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

        private string _fontName;
        public string FontName
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

        private string _outputDocxPath;
        public string OutputDocxPath
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

        private string _inputXlsxPath;
        public string InputXlsxPath
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

        //存储字体的集合
        private List<string> _fonts;
        public List<string> Fonts
        {
            get { return _fonts; }
            set
            {
                _fonts = value;
                OnPropertyChanged(nameof(Fonts));
            }
        }

        //获取系统字体并存至集合fonts中
        public void GetFonts()
        {
            // 创建InstalledFontCollection对象
            InstalledFontCollection installedFonts = new InstalledFontCollection();

            // 获取系统中已安装的字体
            List<System.Drawing.FontFamily> fontFamilies = installedFonts.Families.ToList();
            foreach(System.Drawing.FontFamily family in fontFamilies)
            {
                _fonts.Add(family.Name.ToString());
            }
        }
    }
}