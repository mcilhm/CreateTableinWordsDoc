using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateTableinWordsDoc
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("document.docx");
            FindReplaceOptions opts = new FindReplaceOptions();
            opts.Direction = FindReplaceDirection.Backward;
            opts.ReplacingCallback = new ReplaceEvaluator();

            doc.Range.Replace("[BIKIN_TABLE]", "", opts);
            // Save the Word document
            doc.Save("Find-And-Replace-Text.docx");
        }

        private static Run SplitRun(Run run, int position)
        {
            Run afterRun = (Run)run.Clone(true);
            afterRun.Text = run.Text.Substring(position);
            run.Text = run.Text.Substring((0), (0) + (position));
            run.ParentNode.InsertAfter(afterRun, run);
            return afterRun;
        }

        private class ReplaceEvaluator : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                // ini buat inisialisasi awal atau pun akhir.
                Node currentNode = e.MatchNode;

                // proses pertama (ataupun cuman 1) yang berisi teks replace sebelum disesuaikan,
                // ini untuk berjalan sekaligus memecah replacenya.
                if (e.MatchOffset > 0)
                    currentNode = SplitRun((Run)currentNode, e.MatchOffset);

                // ini buat nampung data yang nantinya buat di replace dan auto remove jika sudah selesai.
                ArrayList runs = new ArrayList();

                // nemuin semua kata yang ingin di replace.
                int remainingLength = e.Match.Value.Length;
                while (
                    (remainingLength > 0) &&
                    (currentNode != null) &&
                    (currentNode.GetText().Length <= remainingLength))
                {
                    runs.Add(currentNode);
                    remainingLength = remainingLength - currentNode.GetText().Length;

                    // setelah di ditemuin tampung data selanjutnya.
                    // harus diperulangin lagi untuk nampung sementara.
                    do
                    {
                        currentNode = currentNode.NextSibling;
                    }
                    while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));
                }

                // Pisahkan proses terakhir yang sesuai jika ada teks yang tersisa.
                if ((currentNode != null) && (remainingLength > 0))
                {
                    SplitRun((Run)currentNode, remainingLength);
                    runs.Add(currentNode);
                }

                //// to insert Table
                DocumentBuilder builder = new
                DocumentBuilder((Document)e.MatchNode.Document);

                builder.MoveTo((Run)runs[runs.Count - 1]);

                Table table = builder.StartTable();

                // untuk menyiapkan data yang akan direplace
                for (int i = 1; i <= 5; i++)
                {
                    for (int j = 1; j <= 5; j++)
                    {

                        builder.InsertCell();
                        builder.Write(string.Format("This is row {0} cell {1}", i, j));
                    }
                    builder.EndRow();
                }

                builder.EndTable();

                foreach (Run run in runs)
                    run.Remove();

                return ReplaceAction.Skip;
            }
        }
    }
}
