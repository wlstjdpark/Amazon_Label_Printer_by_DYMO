using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Media;

namespace LabelsPrint
{
    public static class LabelManager
    {        
        static readonly int rowCount = 4;
        static readonly int columnCount = 10;
        static readonly int density = 1000;
        static Dictionary<int, Dictionary<int, List<LabelItem>>> labelItems = new Dictionary<int, Dictionary<int, List<LabelItem>>>();

        public static string SelectedPrinter { get; set; }        
        public static event Action<string, Color> OnUpdateStatusText;
        public static event Action OnSuccess;
        public static event Action OnFail;        

        class LabelItem
        {            
            string name;
            public string Name { get { return name; } set { name = value; } }
            ImageMagick.MagickImage image;
            DYMO.Label.Framework.ILabel label;

            public LabelItem(ImageMagick.MagickImage image, DYMO.Label.Framework.ILabel label)
            {
                this.image = image;
                this.label = label;
            }

            public void Print(string printerName)
            {
                label.Print(printerName);
            }

            public void WriteToPng()
            {
                image.Write(name + ".png");
            }

            public void WriteToLabel()
            {
                label.SaveToFile(name + ".label");
            }
        }        

        public static void Start(object obj)
        {
            Setting setting = obj as Setting;
            if (obj == null)
            {
                return;
            }

            try
            {
                FileLoad(setting.FileName);
                Process(setting.Option);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "error", MessageBoxButton.OK);
                OnFail?.Invoke();
                return;
            }

            OnSuccess?.Invoke();            
        }        

        private static void FileLoad(string fileName)
        {            
            labelItems.Clear();
            using (var images = new ImageMagick.MagickImageCollection())
            {
                var settings = new ImageMagick.MagickReadSettings();
                settings.Density = new ImageMagick.Density(density, density);

                images.Read(fileName, settings);

                // split
                int page = 1;
                foreach (var image in images)
                {
                    int width = (int)(image.Width / rowCount);
                    int height = (int)(image.Height / columnCount);

                    int number = 0;
                    foreach (var labelImage in image.CropToTiles(width, height))
                    {
                        number++;
                        if (labelImage.TotalColors == 1)
                        {
                            continue;
                        }

                        var encoder = new System.Windows.Media.Imaging.JpegBitmapEncoder();
                        encoder.QualityLevel = 100;

                        using (MemoryStream stream = new MemoryStream())
                        {
                            encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(labelImage.ToBitmapSource()));
                            encoder.Save(stream);
                            stream.Position = 0;

                            var label = DYMO.Label.Framework.Label.Open(
                            Application.GetResourceStream(
                                new Uri("BarcodeAsImage.label", UriKind.RelativeOrAbsolute)).Stream);

                            label.SetImagePngData("Image", stream);

                            var labelItem = new LabelItem(labelImage, label);
                            AddLabelItem(page, number, labelItem);
                        }
                    }
                    page++;
                }
            }
        }

        private static void Process(Setting.ProcessingOption option)
        {
            if ((option & Setting.ProcessingOption.WriteToPng) == Setting.ProcessingOption.WriteToPng)
            {
                OnUpdateStatusText("Writing to png..", Colors.RoyalBlue);
                LabelManager.WriteToPng();
            }

            if ((option & Setting.ProcessingOption.WriteToLabel) == Setting.ProcessingOption.WriteToLabel)
            {
                OnUpdateStatusText("Writing to label..", Colors.RoyalBlue);
                LabelManager.WriteToLabel();
            }

            if ((option & Setting.ProcessingOption.Print) == Setting.ProcessingOption.Print)
            {
                OnUpdateStatusText("Priting..", Colors.RoyalBlue);
                LabelManager.Print();
            }
        }

        private static void AddLabelItem(int page, int number, LabelItem labelItem)
        {
            int pageIndex = page - 1;
            int columnIndex = (number - 1) % rowCount;

            if (labelItems.ContainsKey(pageIndex) == false)
            {
                labelItems.Add(pageIndex, new Dictionary<int, List<LabelItem>>());
            }

            if (labelItems[pageIndex].ContainsKey(columnIndex) == false)
            {
                labelItems[pageIndex].Add(columnIndex, new List<LabelItem>());
            }

            int count = labelItems[pageIndex][columnIndex].Count;
            labelItem.Name = string.Format("{0}_{1}_{2}", pageIndex, columnIndex, count);

            labelItems[pageIndex][columnIndex].Add(labelItem);
        }

        private static IEnumerable<LabelItem> GetOrderedLabelItems()
        {
            for (int i = 0; i < labelItems.Count; ++i)
            {
                for (int j = 0; j < labelItems[i].Count; ++j)
                {
                    for (int k = 0; k < labelItems[i][j].Count; ++k)
                    {
                        yield return labelItems[i][j][k];
                    }
                }
            }
        }

        private static void Print()
        {
            foreach (var labelItem in GetOrderedLabelItems())
            {
                labelItem.Print(SelectedPrinter);
            }            
        }

        private static void WriteToPng()
        {
            foreach (var labelItem in GetOrderedLabelItems())
            {
                labelItem.WriteToPng();
            }            
        }

        private static void WriteToLabel()
        {
            foreach (var labelItem in GetOrderedLabelItems())
            {
                labelItem.WriteToLabel();
            }
        }
    }
}
