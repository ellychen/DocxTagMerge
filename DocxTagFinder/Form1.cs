using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;

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
                MessageBox.Show("來源目錄不存在");
            }
            if (!System.IO.Directory.Exists(txtSavePath.Text.Trim()))
            {
                //建立目的資料夾
                Directory.CreateDirectory(txtSavePath.Text.Trim());
            }

            //來源 Docx 清單
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
                            ErrorMessage += $"{FileName}檔案格式錯誤" + Environment.NewLine;
                            continue;
                        }
                        //                        
                        string SpecialText = "#";
                        List<OpenXmlElement> Result = FindWordPressText(BodyDocument, StartTag);
                        //移除結構正常的標籤節點
                        for (int i = Result.Count - 1; i >= 0; i--)
                        {
                            var ele = Result[i];
                            if (RegexTag.IsMatch(ele.InnerText))
                            {
                                //結構正常 , 不用處理
                                Result.RemoveAt(i);
                            }
                        }


                        foreach (var ele in Result)
                        {
                            if (ele.InnerText == "") continue;
                            //從父節點找出符合的節點
                            var UpEle = ele.Parent;

                            while (UpEle != null && !RegexTag.IsMatch(UpEle.InnerText))
                                UpEle = UpEle.Parent;

                            if (UpEle != null)
                            {
                                //父節點所有 Text
                                var ThisWordTextEles = FindWordPressText(UpEle);
                                //找出開始與結束的節點
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

            MessageBox.Show("處理完成" + (ErrorMessage != "" ? Environment.NewLine : "") + ErrorMessage);
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
        /// 找出符合的節點
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
                    //有符合開頭第一個字元
                    if (ThisElement.Text.IndexOf(StartTag) > -1)
                    {
                        if (ThisElement.Text.IndexOf(PrefixTag) > -1)
                        {
                            //完整開頭標籤
                            IsBegin = true;
                            PassElements.Add(ThisElement);
                            BeginElement = ThisElement;
                        }
                        else
                        {
                            //往下找
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
                                        //不是標準標籤
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
                    // 清除路徑上的 TEXT
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
        /// 標籤文字合併到第一個 Text
        /// </summary>
        /// <param name="Elements"></param>
        private void MergeAccordElement(List<OpenXmlElement> Elements)
        {
            //名稱重組, 放到第一個 wt , 其它清除空白
            string TagName = "";
            for (int i = Elements.Count - 1; i >= 0; i--)
            {
                TagName = Elements[i].InnerText + TagName;
                if (i == 0) continue;
                ((Text)Elements[i]).Text = "";

            }

            ((Text)Elements[0]).Text = TagName;

        }


    }
}