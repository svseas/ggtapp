using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Bold = DocumentFormat.OpenXml.Wordprocessing.Bold;
using Italic = DocumentFormat.OpenXml.Wordprocessing.Italic;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Underline = DocumentFormat.OpenXml.Wordprocessing.Underline;

namespace GGT
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private Run CreateUnderlinedRun(string text)
        {
            Run run = new Run();
            RunProperties runProps = new RunProperties();
            Underline underline = new Underline() { Val = UnderlineValues.Single };
            runProps.Append(underline);
            run.Append(runProps);
            run.Append(new Text(text));
            return run;
        }

        private Run CreateFormattedRun(string content, bool isBold = false, bool isItalic = false)
        {
            Run run = new Run();
            RunProperties runProps = new RunProperties();

            if (isBold)
            {
                runProps.Append(new Bold());
            }

            if (isItalic)
            {
                runProps.Append(new Italic());
            }

            run.Append(runProps);
            run.Append(new Text(content));
            return run;
        }

        private void GenerateDocument(string outputPath,
                              string recipientName,
                              string titlePosition,
                              string organizationUnit,
                              string address,
                              string regarding,
                              string request,
                              DateTime? validityDate,
                              DateTime? createDate)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Insert the Right Box and Left Box
                Table table = body.AppendChild(new Table());

                TableRow row1 = table.AppendChild(new TableRow());

                TableCell cell1 = row1.AppendChild(new TableCell());
                cell1.AppendChild(new Paragraph(new Run(new Text("TÒA ÁN NHÂN DÂN TỐI CAO"))));

                TableCell cell2 = row1.AppendChild(new TableCell());
                cell2.AppendChild(new Paragraph(new Run(new Text("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"))));

                TableRow row2 = table.AppendChild(new TableRow());

                TableCell cell3 = row2.AppendChild(new TableCell());
                cell3.AppendChild(new Paragraph(CreateUnderlinedRun("BÁO CÔNG LÝ")));

                TableCell cell4 = row2.AppendChild(new TableCell());
                cell4.AppendChild(new Paragraph(CreateUnderlinedRun("Độc lập – Tự do – Hạnh phúc")));

                // Insert "GIẤY GIỚI THIỆU"
                Paragraph headerPara = body.AppendChild(new Paragraph());
                Run headerRun = headerPara.AppendChild(new Run());
                headerRun.AppendChild(new Text("GIẤY GIỚI THIỆU"));

                // Insert Recipient Name
                Paragraph recipientNamePara = body.AppendChild(new Paragraph());
                Run recipientNameRun = recipientNamePara.AppendChild(new Run());
                recipientNameRun.AppendChild(new Text($"Đồng chí: {recipientName}"));

                // Insert Title/Position
                Paragraph titlePositionPara = body.AppendChild(new Paragraph());
                Run titlePositionRun = titlePositionPara.AppendChild(new Run());
                titlePositionRun.AppendChild(new Text($"Chức vụ: {titlePosition}"));

                // Insert Organization Unit
                Paragraph organizationUnitPara = body.AppendChild(new Paragraph());
                Run organizationUnitRun = organizationUnitPara.AppendChild(new Run());
                organizationUnitRun.AppendChild(new Text($"Thuộc đơn vị: {organizationUnit}"));

                // Insert Address
                Paragraph addressPara = body.AppendChild(new Paragraph());
                Run addressRun = addressPara.AppendChild(new Run());
                addressRun.AppendChild(new Text($"Được cử đến: {address}"));

                // Insert Regarding
                Paragraph regardingPara = body.AppendChild(new Paragraph());
                Run regardingRun = regardingPara.AppendChild(new Run());
                regardingRun.AppendChild(new Text($"Về việc: {regarding}"));

                // Insert Request
                Paragraph requestPara = body.AppendChild(new Paragraph());
                Run requestRun = requestPara.AppendChild(new Run());
                requestRun.AppendChild(new Text($"Đề nghị Quý cơ quan giúp đỡ để đồng chí hoàn thành nhiệm vụ: {request}"));

                // Insert Validity Date
                if (validityDate.HasValue)
                {
                    Paragraph validityDatePara = body.AppendChild(new Paragraph());
                    Run validityDateRun = validityDatePara.AppendChild(new Run());
                    validityDateRun.AppendChild(new Text($"Giấy Giới thiệu có giá trị đến ngày: {validityDate.Value:dd/MM/yyyy}"));
                }

                // Insert footer with Create Date, position, name, and phone number
                Paragraph createDatePara = body.AppendChild(new Paragraph());
                createDatePara.AppendChild(CreateFormattedRun($"Hà Nội, {createDate.Value:dd/MM/yyyy}", isItalic: true));

                Paragraph positionPara = body.AppendChild(new Paragraph());
                positionPara.AppendChild(CreateFormattedRun("TỔNG BIÊN TẬP", isBold: true));

                Paragraph namePara = body.AppendChild(new Paragraph());
                namePara.AppendChild(CreateFormattedRun("Trần Đức Vinh"));

                Paragraph phonePara = body.AppendChild(new Paragraph());
                phonePara.AppendChild(CreateFormattedRun("Điện thoại: 024.3824.7204 – 024. 3936.5550", isItalic: true));

                // Save the changes to the main document part
                mainPart.Document.Save();
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Get values from WPF input fields
            string recipientName = RecipientNameTextBox.Text;
            string titlePosition = TitlePositionTextBox.Text;
            string organizationUnit = OrganizationUnitTextBox.Text;
            string address = AddressTextBox.Text;
            string regarding = RegardingTextBox.Text;
            string request = RequestTextBox.Text;
            DateTime? validityDate = ValidityDateDatePicker.SelectedDate;
            DateTime? createDate = CreateDateDatePicker.SelectedDate;

            // Call GenerateDocument with the values
            GenerateDocument("C:\\Users\\ThinkPad\\Documents\\Congly\\GGT.docx", recipientName, titlePosition, organizationUnit, address, regarding, request, validityDate, createDate);
        }

    }
}