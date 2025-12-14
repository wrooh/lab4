using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

namespace WpfApp2.HelperClasses
{
    // Класс для работы с документами Word. 
    public class WordHelper
    {
        private FileInfo _fileInfo;

        public WordHelper(string filename)
        {
            if (File.Exists(filename))
            {
                _fileInfo = new FileInfo(filename);
            }
            else
            {
                throw new FileNotFoundException();
            }
        }

        public void Process(Dictionary<string, string> items, string path)
        {
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                object file = _fileInfo.FullName;
                object missing = Type.Missing;

                app.Documents.Open(file);

                foreach (var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    object wrap = Word.WdFindWrap.wdFindContinue;
                    object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(
                        FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing,
                        Replace: replace);
                }

                app.ActiveDocument.SaveAs2(path);
                app.ActiveDocument.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                }
            }
        }
    }
}