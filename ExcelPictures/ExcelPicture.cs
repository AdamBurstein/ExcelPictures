using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using OfficeOpenXml;

namespace ExcelPictures
{
    class ExcelPicture
    {
        public string[] clArgs { get; set; }
        private Dictionary<string, string> clArgsDict;
        private List<List<String>> redTrack;
        private List<List<String>> greenTrack;
        private List<List<String>> blueTrack;
        ExcelPackage package;
        ExcelWorksheet myWorksheet;
        private int numCols;
        private int numRows;


        public ExcelPicture()
        {
            clArgs = new string[] { };
            clArgsDict = new Dictionary<string, string>();
            redTrack = new List<List<string>>();
            greenTrack = new List<List<string>>();
            blueTrack = new List<List<string>>();
            package = new ExcelPackage();
            myWorksheet = package.Workbook.Worksheets.Add("Picture");
        }

        public void begin()
        {
            DateTime startTime = DateTime.Now;
            Console.WriteLine("Started at: " + startTime.ToString("HH:mm:ss tt"));
            foreach (String str in clArgs)
            {
                String[] nameValue = str.Split(new char[] { '=' });
                clArgsDict.Add(nameValue[0], nameValue[1]);
            }

            getInputFile();
            getOutputFile();
            getScale();
            parseFile();
            writeFile();
            package.SaveAs(new FileInfo(clArgsDict["out"]));
            DateTime endTime = DateTime.Now;
            Console.WriteLine("Ended at: " + endTime.ToString("HH:mm:ss tt"));
            TimeSpan elapsedTime = endTime.Subtract(startTime);
            Console.WriteLine("Elapsed: " + elapsedTime.ToString());
        }

        private void getInputFile()
        {
            Boolean fileExists = false;

            if (!clArgsDict.ContainsKey("in"))
                fileExists = false;
            else
                fileExists = System.IO.File.Exists(clArgsDict["in"]);

            while (!fileExists)
            { 
                String inputFile = Console.ReadLine();
                inputFile = inputFile.Replace("~", "/users/" + Environment.GetEnvironmentVariable("USER"));
                if (clArgsDict.ContainsKey("in"))
                    clArgsDict["in"] = inputFile;
                else
                    clArgsDict.Add("in", inputFile);
                fileExists = System.IO.File.Exists(clArgsDict["in"]);
            }


        }

        private void writeRedLine(List<String> line, int row)
        {
            int col = 1;
            int writeRow = (3*row) + 1;
            foreach (String str in line)
            {
                ExcelRange theRange = myWorksheet.Cells[writeRow, col];
                theRange.Value = str;
                int color = Int32.Parse(str);
                theRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                theRange.Style.Fill.BackgroundColor.SetColor(1, color, 0, 0);
                theRange.Style.Font.Color.SetColor(1, color, 0, 0);
                ++col;
            }
        }

        private void writeGreenLine(List<String> line, int row)
        {
            int col = 1;
            int writeRow = (3 * row) + 2;
            foreach (String str in line)
            {
                ExcelRange theRange = myWorksheet.Cells[writeRow, col];
                theRange.Value = str;
                int color = Int32.Parse(str);
                theRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                theRange.Style.Fill.BackgroundColor.SetColor(1, 0, color, 0);
                theRange.Style.Font.Color.SetColor(1, 0, color, 0);
                ++col;
            }
        }

        private void writeBlueLine(List<String> line, int row)
        {
            int col = 1;
            int writeRow = (3 * row) + 3;
            foreach (String str in line)
            {
                ExcelRange theRange = myWorksheet.Cells[writeRow, col];
                theRange.Value = str;
                int color = Int32.Parse(str);
                theRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                theRange.Style.Fill.BackgroundColor.SetColor(1, 0, 0, color);
                theRange.Style.Font.Color.SetColor(1, 0, 0, color);
                ++col;
            }
        }

        private void writeFile()
        {
            int numLists = redTrack.Count;
            int numPixelsWritten = 0;

            for (int i = 0; i < numLists; ++i)
            {
                Console.WriteLine("Writing file: " + ++numPixelsWritten + " of " + numLists + " lists written");
                writeRedLine(redTrack[i], i);
                writeGreenLine(greenTrack[i], i);
                writeBlueLine(blueTrack[i], i);
            }

            for (int i = 0; i < numRows * 3; ++i)
            {
                myWorksheet.Row(i + 1).Height = 7;
            }
            for (int i = 0; i < numCols; ++i)
            {
                myWorksheet.Column(i+1).Width = 4;
            }
            myWorksheet.View.ZoomScale = 10;
        }

        private void parseFile()
        {

            Bitmap image = new Bitmap(clArgsDict["in"]);
            numRows = image.Height;
            numCols = image.Width;
            int scale = Int32.Parse(clArgsDict["scale"]);
            int numRowsOutput = image.Height / scale;
            int numColsOutput = image.Width / scale;

            long totalPixelsToAnalyze = numRowsOutput * numColsOutput;
            long totalPixelsAnalyzed = 0;
            long onePercent = totalPixelsToAnalyze / 100;

            Console.WriteLine("Image dimensions: " + numRows + "H x " + numCols + "W");

            for (int row = 0; row < numRows; row += scale)
            {
                List<String> redLine = new List<String>();
                List<String> greenLine = new List<String>();
                List<String> blueLine = new List<String>();

                for (int col = 0; col < numCols; col += scale)
                {
                    totalPixelsAnalyzed++;
                    Console.WriteLine("Parsing: " + totalPixelsAnalyzed + "/" + totalPixelsToAnalyze);
                    try
                    {
                        Color pixelInfo = image.GetPixel(col, row);

                        redLine.Add(pixelInfo.R.ToString());
                        greenLine.Add(pixelInfo.G.ToString());
                        blueLine.Add(pixelInfo.B.ToString());
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error: " + e.Message);
                    }
                }

                redTrack.Add(redLine);
                greenTrack.Add(greenLine);
                blueTrack.Add(blueLine);
            }

            Console.WriteLine();
        }

        private void getOutputFile()
        {
            if (clArgsDict.ContainsKey("out"))
            {
                clArgsDict["out"] = clArgsDict["out"].Replace("~", "/users/" + Environment.GetEnvironmentVariable("USER"));
                return;
            }
            Console.Write("Please enter the path to the output Excel file: ");
            String outputFile = Console.ReadLine();
            outputFile = outputFile.Replace("~", "/users/" + Environment.GetEnvironmentVariable("USER"));
            clArgsDict.Add("out", outputFile);

        }

        private void getScale()
        {
            if (clArgsDict.ContainsKey("scale"))
                return;
            Console.Write("Scale image by: ");
            String scale = Console.ReadLine();
            clArgsDict.Add("scale", scale);
        }

        static void Main(string[] args)
        {
            ExcelPicture p = new ExcelPicture();
            p.clArgs = args;
            p.begin();

        }
    }
}
