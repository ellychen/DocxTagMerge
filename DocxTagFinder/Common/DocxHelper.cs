

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public class DocxHelper
{



    private WordprocessingDocument Word;

    private Dictionary<string, OpenXmlElement> Bookmarks;
    public DocxHelper(string FilePath)
    {
        Word = WordprocessingDocument.Open(FilePath, true);
        FilterAllBookmarks();
    }

    public DocxHelper(System.IO.Stream stream)
    {
        Word = WordprocessingDocument.Open(stream, true);
        FilterAllBookmarks();
    }


    private void FilterAllBookmarks()
    {
        if (Word == null) return;


        Bookmarks = LoadBookmarkElement(Word.MainDocumentPart.Document)
                            .ToDictionary(x => x.GetAttributes().FirstOrDefault(attr => attr.LocalName.ToLower() == "name").Value, x => x);
        /*
        Bookmarks = Word.MainDocumentPart.Document
                               .Where(d => d.LocalName.ToLower() == "bookmarkstart" && d.GetAttributes().Any(attr => attr.LocalName.ToLower() == "name"))
                               .ToDictionary(x => x.GetAttributes().FirstOrDefault(attr => attr.LocalName.ToLower() == "name").Value, x => x);
        */


    }


    private List<OpenXmlElement> LoadBookmarkElement(OpenXmlElement Element)
    {
        List<OpenXmlElement> list = new List<OpenXmlElement>();
        if (Element.HasChildren)
        {
            foreach (var c in Element.ChildElements)
            {
                list.AddRange(LoadBookmarkElement(c));
            }
        }

        if (Element.LocalName.ToLower() == "bookmarkstart" && Element.GetAttributes().Any(x => x.LocalName.ToLower() == "name"))
            list.Add(Element);


        return list;
    }




    public bool AddProtect(string ProtectPD)
    {
        /*
              * 1. Document is password protected. (文件密碼保護)
              * 2. Document is recommended to be opened as read-only. (建議唯讀) 
              * 4. Document is enforced to be opened as read-only. (強制唯讀)
              * 8. Document is locked for annotation. (文件被鎖定)
              * 
              */
        Word.ExtendedFilePropertiesPart.Properties.DocumentSecurity =
            new DocumentFormat.OpenXml.ExtendedProperties.DocumentSecurity("8");
        Word.ExtendedFilePropertiesPart.Properties.Save();


        var trackRevisions = new TrackRevisions();

        var dp = new DocumentProtection();

        // 編輯模式 , 追踪修訂
        dp.Edit = DocumentProtectionValues.TrackedChanges;
        dp.Enforcement = OnOffValue.FromBoolean(true);

        // 密碼設定
        dp.Hash = new Base64BinaryValue(ProtectPD);

        // 強制

        Word.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(dp);
        Word.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(new TrackRevisions());


        return true;
    }

    /// <summary>
    /// 移除追踪
    /// </summary>
    /// <returns></returns>
    public bool RemoveProtect()
    {
        bool Result = false;
        //移除 Settings.Xml protected setting node 
        RemoveProtectSettings();
        //移除 Change Log Node 
        RemoveDocumentChangeTag();
        //移除 Delete Log Node 
        RemoveDocumentDeleteTag();

        return Result;

    }

    private void RemoveProtectSettings()
    {
        var dps = Word.MainDocumentPart.DocumentSettingsPart.Settings.Where(n => n.LocalName.ToLower() == "documentprotection").ToList();
        var trs = Word.MainDocumentPart.DocumentSettingsPart.Settings.Where(n => n.LocalName.ToLower() == "trackrevisions").ToList();

        if (dps.Any())
        {
            foreach (var dp in dps) Word.MainDocumentPart.DocumentSettingsPart.Settings.RemoveChild(dp);
        }

        if (trs.Any())
        {
            foreach (var tr in trs) Word.MainDocumentPart.DocumentSettingsPart.Settings.RemoveChild(tr);
        }
    }

    private void RemoveDocumentChangeTag()
    {

        List<OpenXmlElement> changes =
                Word.MainDocumentPart.Document.Body.Descendants<ParagraphPropertiesChange>()
                .Where(c => c.Author.Value != "").Cast<OpenXmlElement>().ToList();

        foreach (OpenXmlElement change in changes)
        {
            change.Remove();
        }
    }

    private void RemoveDocumentDeleteTag()
    {
        List<OpenXmlElement> deletions =
                     Word.MainDocumentPart.Document.Body.Descendants<Deleted>()
                    .Where(c => c.Author.Value != "").Cast<OpenXmlElement>().ToList();

        deletions.AddRange(Word.MainDocumentPart.Document.Body.Descendants<DeletedRun>()
            .Where(c => c.Author.Value != "").Cast<OpenXmlElement>().ToList());

        deletions.AddRange(Word.MainDocumentPart.Document.Body.Descendants<DeletedMathControl>()
            .Where(c => c.Author.Value != "").Cast<OpenXmlElement>().ToList());

        foreach (OpenXmlElement deletion in deletions)
        {
            deletion.Remove();
        }

    }



    public bool ChangeBookmarkValue(string BookmarkName, string val)
    {
        bool HasRun = false;
        if (Bookmarks == null) return false;
        var MapBM = Bookmarks.Where(d => d.Key == BookmarkName).Select(d => d.Value).ToList();
        if (!MapBM.Any()) return false;
        foreach (var ele in MapBM)
        {
            //第一個 w:r , 取代其它的內容
            var ThisId = ele.GetAttributes().FirstOrDefault(at => at.LocalName.ToLower() == "id").Value;
            // 找到 bookmarkEnd 所有 element

            var NearElement = ele.NextSibling();
            //不能指向自己            
            if (NearElement == ele) NearElement = NearElement.NextSibling();
            //最終只要留第一個 , 其它刪除
            List<OpenXmlElement> SiblingElements = new List<OpenXmlElement>();
            while (NearElement != null)
            {
                if (NearElement.LocalName.ToLower() == "bookmarkend" &&
                    NearElement.GetAttributes().FirstOrDefault(at => at.LocalName.ToLower() == "id").Value == ThisId)
                {
                    break;
                }
                else
                {

                    SiblingElements.Add(NearElement);
                }

                NearElement = NearElement.NextSibling();
            }


            var write = SiblingElements.FirstOrDefault(x => x.LocalName == "r") ?? new Run();
            // ele.ChildElements.FirstOrDefault(x => x.LocalName == "r");
            //
            if (write == null) continue;

            //找出 w:t
            if (val.Length == 4 && val.IndexOf("F0") == 0)
            {
                var SymArea = write.ChildElements.FirstOrDefault(x => x.LocalName == "sym");
                if (SymArea != null) SymArea.Remove();
                write.InnerXml += write.InnerXml + "<w:sym w:font=\"Wingdings\" w:char=\"" + val + "\" />";
            }
            else
            {
                var TextArea = write.ChildElements.FirstOrDefault(x => x.LocalName == "t");
                if (TextArea == null)
                {
                    write.AppendChild(new Text(val));
                }
                else
                {
                    ((Text)TextArea).Text = val;
                    //TextArea.InnerXml = val;
                }
            }

            var NewWrite = write.CloneNode(true);
            //var NewContent = write.InnerXml; //.ToString();
            foreach (var se in SiblingElements) se.Remove();

            ele.InsertAfterSelf(NewWrite);

            HasRun = true;
        }

        return HasRun;

        /*
              <w:r>
                <w:rPr>
                  <w:rFonts w:ascii="標楷體" w:eastAsia="標楷體" w:hint="eastAsia"/>
                  <w:b/>
                  <w:sz w:val="32"/>
                  <w:shd w:val="pct15" w:color="auto" w:fill="FFFFFF"/>
                </w:rPr>
                <w:t>pname</w:t>
              </w:r>
        */
    }


    public void Close()
    {
        Word.Close();
        Word.Dispose();

    }
    public void CloseAsSave()
    {
        //Word.Save();

    }

}


