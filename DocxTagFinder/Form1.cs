using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using odf = Independentsoft.Office.Odf;

namespace DocxTagFinder
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            txtRawFile.Text = Environment.CurrentDirectory;
            txtSavePath.Text = Environment.CurrentDirectory;
        }

        private Regex RegexTag;
        private string ShortTag = "#";
        private string fixTag = "##";
        private void btnFinderClick(object sender, EventArgs e)
        {

            if (!System.IO.Directory.Exists(txtRawFile.Text.Trim()))
            {
                MessageBox.Show("�ӷ��ؿ����s�b");
            }
            if (!System.IO.Directory.Exists(txtSavePath.Text.Trim()))
            {
                //�إߥت���Ƨ�
                Directory.CreateDirectory(txtSavePath.Text.Trim());
            }

            //�ӷ� Docx �M��
            List<string> DocxFiles = new List<string>();
            foreach (var docx in System.IO.Directory.GetFiles(txtRawFile.Text.Trim()))
            {
                if (System.IO.Path.GetExtension(docx).ToLower() == ".docx")
                {
                    DocxFiles.Add(docx);
                }
            }

            //List<string> DocxFiles = new List<string>();
            //DocxFiles.Add(@"D:\SideProject\DocxTagFinder\DocxTagFinder\05.docx");



            string StartTag = txtStartTag.Text.Trim();
            string EndTag = txtEndTag.Text.Trim();
            string PrefixTag = txtPrefixTag.Text.Trim();
            string CloseTag = txtCloseTag.Text.Trim();

            RegexTag = new Regex(@$"({PrefixTag})+(\w+)+({CloseTag})");


            string ErrorMessage = "";

            foreach (var FilePath in DocxFiles)
            {
                string FileName = Path.GetFileNameWithoutExtension(FilePath);
                string SnapPath = System.IO.Path.Combine(txtSavePath.Text.Trim()
                                            , $"{FileName}_{DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss")}{Path.GetExtension(FilePath)}");

                System.IO.File.Copy(FilePath, SnapPath, true);
                //
                try
                {
                    using (WordprocessingDocument Word = WordprocessingDocument.Open(SnapPath, true))
                    {
                        if (Word == null) continue;
                        //
                        var BodyDocument = Word?.MainDocumentPart?.Document;
                        if (BodyDocument == null)
                        {
                            ErrorMessage += $"{FileName}�ɮ׮榡���~" + Environment.NewLine;
                            continue;
                        }
                        //                        
                        string SpecialText = "#";
                        List<OpenXmlElement> Result = FindWordPressText(BodyDocument, StartTag);
                        //�������c���`�����Ҹ`�I
                        for (int i = Result.Count - 1; i >= 0; i--)
                        {
                            var ele = Result[i];
                            if (RegexTag.IsMatch(ele.InnerText))
                            {
                                //���c���` , ���γB�z
                                Result.RemoveAt(i);
                            }
                        }


                        foreach (var ele in Result)
                        {
                            if (ele.InnerText == "") continue;
                            //�q���`�I��X�ŦX���`�I
                            var UpEle = ele.Parent;

                            while (UpEle != null && !RegexTag.IsMatch(UpEle.InnerText))
                                UpEle = UpEle.Parent;

                            if (UpEle != null)
                            {
                                //���`�I�Ҧ� Text
                                var ThisWordTextEles = FindWordPressText(UpEle);
                                //��X�}�l�P�������`�I
                                MergeIntervalElement(ThisWordTextEles, StartTag, EndTag, PrefixTag, CloseTag);
                            }
                        }

                        Word.Save();
                        Word.Close();
                    }
                }
                catch (Exception ex)
                {
                    ErrorMessage += $"{FileName} : {ex.Message}" + Environment.NewLine;
                }
            }

            MessageBox.Show("�B�z����" + (ErrorMessage != "" ? Environment.NewLine : "") + ErrorMessage);
        }


        private List<OpenXmlElement> FindWordPressText(OpenXmlElement content, string SpecialText = "")
        {
            List<OpenXmlElement> list = new List<OpenXmlElement>();

            if (content.HasChildren)
            {
                foreach (var Child in content.ChildElements)
                {
                    var child_list = FindWordPressText(Child, SpecialText);
                    if (child_list != null && child_list.Count > 0)
                        list.AddRange(child_list);
                }

            }
            else
            {
                if (content is Text)
                {
                    if (string.IsNullOrWhiteSpace(SpecialText))
                    {
                        list.Add(content);
                    }
                    else
                    {
                        if (content.InnerText.IndexOf(SpecialText) > -1)
                        {
                            list.Add(content);
                        }
                    }
                }
            }

            return list;

        }

        /// <summary>
        /// ��X�ŦX���`�I
        /// </summary>
        /// <param name="ParentElements"></param>
        /// <param name="RawElements"></param>
        /// <param name="TagKey"></param>
        /// <returns></returns>
        private void MergeIntervalElement(List<OpenXmlElement> ParentElements, string StartTag, string EndTag, string PrefixTag, string CloseTag)
        {
            bool IsBegin = false;
            OpenXmlElement BeginElement = null;
            List<OpenXmlElement> PassElements = new List<OpenXmlElement>();
            int BeginIdx = -1;
            int EndIdx = -1;
            List<OpenXmlElement> ThisElements = new List<OpenXmlElement>();
            for (int i = 0; i < ParentElements.Count; i++)
            {
                Text ThisElement = ParentElements[i] as Text;
                if (ThisElement == null) continue;

                if (!IsBegin)
                {
                    //���ŦX�}�Y�Ĥ@�Ӧr��
                    if (ThisElement.Text.IndexOf(StartTag) > -1)
                    {
                        if (ThisElement.Text.IndexOf(PrefixTag) > -1)
                        {
                            //����}�Y����
                            IsBegin = true;
                            PassElements.Add(ThisElement);
                            BeginElement = ThisElement;
                        }
                        else
                        {
                            //���U��
                            while (true)
                            {
                                ++i;
                                var SiblingElement = ParentElements[i] as Text;
                                if (SiblingElement != null && SiblingElement.Text != "")
                                {
                                    var eleText = ThisElement.Text + SiblingElement.Text;
                                    if (eleText.IndexOf(PrefixTag) > -1)
                                    {
                                        IsBegin = true;
                                        PassElements.Add(ThisElement);
                                        PassElements.Add(SiblingElement);
                                        BeginElement = ThisElement;
                                    }
                                    else
                                    {
                                        //���O�зǼ���
                                        IsBegin = false;
                                    }

                                    break;
                                }
                            }


                        }
                    }
                }
                else
                {
                    PassElements.Add(ThisElement);
                }

                string PassInnerText = string.Join("", PassElements.Select(x => x.InnerText).ToArray());
                if (RegexTag.IsMatch(PassInnerText))
                {
                    ((Text)BeginElement).Text = PassInnerText;
                    // �M�����|�W�� TEXT
                    for (int x = 1; x < PassElements.Count; x++)
                    {
                        var Pass = PassElements[x] as Text;
                        if (Pass != null)
                            Pass.Text = "";
                    }

                    BeginElement = null;
                    IsBegin = false;
                    PassElements = new List<OpenXmlElement>();
                }
            }
        }


        private IEnumerable<OpenXmlElement> ElementYield(List<OpenXmlElement> RawElements)
        {
            while (true)
            {
                if (RawElements.Count == 0 || RawElements == null) break;
                //
                yield return RawElements[0];
            }
        }


        /// <summary>
        /// ���Ҥ�r�X�֨�Ĥ@�� Text
        /// </summary>
        /// <param name="Elements"></param>
        private void MergeAccordElement(List<OpenXmlElement> Elements)
        {
            //�W�٭���, ���Ĥ@�� wt , �䥦�M���ť�
            string TagName = "";
            for (int i = Elements.Count - 1; i >= 0; i--)
            {
                TagName = Elements[i].InnerText + TagName;
                if (i == 0) continue;
                ((Text)Elements[i]).Text = "";

            }

            ((Text)Elements[0]).Text = TagName;

        }





        private void btnRun2_ODFClick(object sender, EventArgs e)
        {


            string StartTag = txtStartTag.Text.Trim();
            string EndTag = txtEndTag.Text.Trim();
            string PrefixTag = txtPrefixTag.Text.Trim();
            string CloseTag = txtCloseTag.Text.Trim();
            RegexTag = new Regex(@$"({PrefixTag})+(\w+)+({CloseTag})");
            //


            if (!System.IO.Directory.Exists(txtRawFile.Text.Trim()))
            {
                MessageBox.Show("�ӷ��ؿ����s�b");
            }
            if (!System.IO.Directory.Exists(txtSavePath.Text.Trim()))
            {
                //�إߥت���Ƨ�
                Directory.CreateDirectory(txtSavePath.Text.Trim());
            }

            //�ӷ� Docx �M��
            List<string> DocxFiles = new List<string>();
            foreach (var docx in System.IO.Directory.GetFiles(txtRawFile.Text.Trim()))
            {
                if (System.IO.Path.GetExtension(docx).ToLower() == ".odt")
                {
                    DocxFiles.Add(docx);
                }
            }

            string ErrorMessage = "";
            foreach (var odt in DocxFiles)
            {
                string FileName = Path.GetFileNameWithoutExtension(odt);
                string SnapPath = System.IO.Path.Combine(txtSavePath.Text.Trim()
                                            , $"{FileName}_{DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss")}{Path.GetExtension(odt)}");

                System.IO.File.Copy(odt, SnapPath, true);
                try
                {
                    ProcessODF(SnapPath, StartTag, PrefixTag);
                }
                catch (Exception ex)
                {
                    ErrorMessage += Path.GetFileName(odt) + " ����  ; " + ex.Message + " \n";
                }

            }
            if (ErrorMessage != "")
                MessageBox.Show(ErrorMessage);

            MessageBox.Show("finish");
        }



        private void ProcessODF(string SnapPath, string StartTag, string PrefixTag)
        {
            odf.TextDocument doc = new odf.TextDocument(SnapPath);
            // ���o�C�@�檺 XML 
            var bodyContents = doc.Body.Content;

            // �C�@�� Element 
            //       �o�̥]�t�F Text (�]�tXML) �� AttributedText (�u�����e��r)
            // bodyContents[0].GetContentElements();

            List<odf.ITextContent> TagLines = new List<odf.ITextContent>();
            foreach (var line in bodyContents)
            {
                //��ڦr�ꤺ�e

                string LineString = string.Join("",
                            line.GetContentElements()
                                        .Where(d => d is odf.Text)
                                        .Select(d => d.ToString()).ToArray()
                            );


                if (RegexTag.IsMatch(LineString))
                {
                    odf.Text BeginElement = null;
                    List<odf.Text> PathElements = new List<odf.Text>();

                    //�ˬd ��@ Text �N�ŦX���� 
                    // isAccept 
                    bool NeedProcess = true;
                    foreach (odf.Text text in line.GetContentElements().Where(d => d is odf.Text))
                    {
                        if (RegexTag.IsMatch(text.Value))
                        {
                            NeedProcess = false;
                            break;
                        }
                    }
                    if (!NeedProcess)
                        continue;
                    //

                    var Elements = line.GetContentElements();
                    for (int i = 0; i < Elements.Count; i++)
                    {
                        var Ele = Elements[i];
                        if (Ele is odf.Text)
                        {
                            var TextEle = Ele as odf.Text;
                            if (TextEle == null) continue;
                            //

                            if (BeginElement != null)
                            {
                                PathElements.Add(TextEle);
                            }
                            else
                            {
                                if (TextEle.ToString().IndexOf(StartTag) > -1)
                                {
                                    // ���_�l�аO ,  �X�֫᭱���r��O�_����ŦX
                                    while (true)
                                    {
                                        ++i;
                                        if (Elements[i] is not odf.Text) continue;
                                        if (Elements[i].ToString() == "") continue;

                                        if ((Ele.ToString() + Elements[i].ToString()).IndexOf(PrefixTag) > -1)
                                        {
                                            //�O���Ұ��I
                                            BeginElement = TextEle;
                                            // 
                                            PathElements.Add(TextEle);
                                            PathElements.Add(Elements[i] as odf.Text);

                                        }
                                        else
                                        {
                                            //����                                        
                                        }

                                        break;
                                    }
                                }
                            }

                            //�ˬd�r��
                            if (RegexTag.IsMatch(
                                    string.Join("", PathElements.Select(t => t.ToString()).ToArray())))
                            {
                                //�r�꧹���ŦX

                                // ��r�������i BeginElement 
                                BeginElement.Value = string.Join("", PathElements.Select(t => t.ToString()).ToArray());
                                // ��PathElement ��Value (���F��0��) �����M��
                                for (int x = 1; x < PathElements.Count; x++)
                                {
                                    PathElements[x].Value = "";
                                }

                                //�����ˬd , �קK���ĤG�� { 
                            }
                        }
                    }
                }
            }

            doc.Save(SnapPath, true);
        }
    }
}