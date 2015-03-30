using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SeniorProject
{
    class CreateStyle
    {
        string style = "";
        string words = "";
        public CreateStyle()
        {
            this.style = "<styles>" +
                          "<style>" +
                            "<Nameformat>วิศวกรรมศาสตร์</Nameformat>" +
                            "<Margin>3.81,2.54,2.54,2.54</Margin>" +
                            "<fonts>" +
                              "<font>" +
                                "<Namefont>Angsana New</Namefont>" +
                                "<TH>" +
                                    "<CoverTitle>20</CoverTitle>" +
                                    "<CoverOperator>18</CoverOperator>" +
                                    "<Chapter>20</Chapter>" +
                                    "<Namechapter>20</Namechapter>" +
                                    "<Topics>18</Topics>" +
                                    "<Subheading>16</Subheading>" +
                                    "<Substance>16</Substance>" +
                                "</TH>" +
                                "<EN>" +
                                    "<CoverTitle>18</CoverTitle>" +
                                    "<CoverOperator>18</CoverOperator>" +
                                    "<Chapter>20</Chapter>" +
                                    "<Namechapter>20</Namechapter>" +
                                    "<Topics>18</Topics>" +
                                    "<Subheading>16</Subheading>" +
                                    "<Substance>16</Substance>" +
                                "</EN>" +
                              "</font>" +
                              "<font>" +
                               " <Namefont>AngsanaUPC</Namefont>" +
                               " <TH>" +
                                   "<CoverTitle>20</CoverTitle>" +
                                   "<CoverOperator>18</CoverOperator>" +
                                    "<Chapter>20</Chapter>" +
                                    "<Namechapter>20</Namechapter>" +
                                    "<Topics>18</Topics>" +
                                    "<Subheading>16</Subheading>" +
                                    "<Substance>16</Substance>" +
                                "</TH>" +
                                "<EN>" +
                                    "<CoverTitle>18</CoverTitle>" +
                                    "<CoverOperator>18</CoverOperator>" +
                                    "<Chapter>20</Chapter>" +
                                    "<Namechapter>20</Namechapter>" +
                                    "<Topics>18</Topics>" +
                                    "<Subheading>16</Subheading>" +
                                    "<Substance>16</Substance>" +
                                "</EN>" +
                              "</font>" +
                            "</fonts>" +
                            "<Paper>wdPaperA4</Paper>" +
                            "<Departments>" +
                                "<Department>วิศวกรรมโยธา เครื่องกล และอุตสาหการ</Department>" +
                                "<Department>วิศวกรรมไฟฟ้าและคอมพิวเตอร์</Department>" +
                            "</Departments>" +
                            "<Indent>36.0</Indent>" +
                          "</style>" +
                          "<style>" +
                            "<Nameformat>บัณฑิตวิทยาลัย มน</Nameformat>" +
                            "<Margin>3.75,2.5,3.75,2.5</Margin>" +
                            "<fonts>" +
                              "<font>" +
                                "<Namefont>Cordia New</Namefont>" +
                                "<TH>" +
                                    "<CoverTitle>16</CoverTitle>" +
                                    "<CoverOperator>16</CoverOperator>" +
                                    "<Chapter>18</Chapter>" +
                                    "<Namechapter>18</Namechapter>" +
                                    "<Topics>18</Topics>" +
                                    "<Subheading>16</Subheading>" +
                                    "<Substance>16</Substance>" +
                                "</TH>" +
                                "<EN>" +
                                    "<CoverTitle>16</CoverTitle>" +
                                    "<CoverOperator>16</CoverOperator>" +
                                    "<Chapter>18</Chapter>" +
                                    "<Namechapter>18</Namechapter>" +
                                    "<Topics>18</Topics>" +
                                    "<Subheading>16</Subheading>" +
                                    "<Substance>16</Substance>" +
                                "</EN>" +
                              "</font>" +
                              "<font>" +
                                "<Namefont>Times New Roman</Namefont>" +
                                "<TH>" +
                                    "<CoverTitle>12</CoverTitle>" +
                                    "<CoverOperator>12</CoverOperator>" +
                                    "<Chapter>18</Chapter>" +
                                    "<Namechapter>18</Namechapter>" +
                                    "<Topics>18</Topics>" +
                                    "<Subheading>16</Subheading>" +
                                    "<Substance>16</Substance>" +
                                "</TH>" +
                                "<EN>" +
                                    "<CoverTitle>12</CoverTitle>" +
                                    "<CoverOperator>12</CoverOperator>" +
                                    "<Chapter>16</Chapter>" +
                                    "<Namechapter>16</Namechapter>" +
                                    "<Topics>12</Topics>" +
                                    "<Subheading>12</Subheading>" +
                                    "<Substance>12</Substance>" +
                                "</EN>" +
                              "</font>" +
                            "</fonts>" +
                            "<Paper>wdPaperA4</Paper>" +
                            "<Departments>" +
                            "</Departments>" +
                            "<Indent>42.52</Indent>" +
                          "</style>" +
                        "</styles>";
            //==============================================//
            this.words = "แมโคร,มาโคร,มาโค\n";
            this.words += "เรจิสเตอร์,รีจีสเตอร์,รีจิสเตอร์\n";
            this.words += "เมทริกซ์,แมทริก,เมตทริก,\n";

            this.words += "เมท็อด,เม็ทตอด,เมทตอด\n";

            this.words += "แมนทิสซา,แมนติสซ่า,แมนทิสซ่า,แมนติทซ่า,แมนทิทซ่า\n";

            this.words += "เบราว์เซอร์,บราวเซอร์\n";

            this.words += "แอ็กทิฟ,แอ๊กทีป,เอทีฟ\n";

            this.words += "เอกซ์ทราเน็ต,แอ๊กทราเน็ต\n";

            this.words += "มอดุลาร์,มอดุลล่า\n";

            this.words += "แอนะล็อก,อนาลอก,อนาล๊อก\n";

            this.words += "เอนทิตี,เอ็นติตี้\n";

            this.words += "แอปเพล็ต,แอปเพล็ท\n";

            this.words += "มัลติ,มันติ\n";

            this.words += "อัลกอล,อันกอล\n";

            this.words += "มิดเดิล,มิเดิล\n";

            this.words += "อาร์กิวเมนต์,อะกิวเม้น\n";

            this.words += "แอสเซมเบลอร์,แอสซิมเบลอร์\n";

            this.words += "เกตเวย์,เก็ตเว\n";

            this.words += "โทเค็น,โทคเค่น\n";

            this.words += "เทลเน็ต,เท็วเน็ตม,เทวเน็ต\n";

            this.words += "แบนด์วิดท์,แบนวิด\n";

            this.words += "เพรดิเคต,พรีดิกเคต\n";

            this.words += "โพรโทคอล,โปรโตคอ\n";

            this.words += "อ็อบเจกต์,อ๊อปเจ๊ก\n";

            this.words += "ซ็อกเก็ต,ซ๊อกเก็ต\n";

            this.words += "ไฟร์วอลล์,ไฟล์วอ\n";

            this.words += "ดิจิทัล,ดิจิตอล\n";

            this.words += "ฟลิปฟล็อป,ฟลิ๊กฟลอป\n";

            this.words += "ลินุกซ์,ลีนุกซ์,ลีนุก\n";

        }

        public void Check(){
            string folder = "CheckingThesis";
            string pathDoc = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            string path = pathDoc + "\\" + folder;
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string pathFile = path + "\\" + "stlyes.xml";
            CreateStyleFile(pathFile);

            pathFile = path + "\\" + "Royal.txt";
            CreateWordFile( pathFile);
            /*if (!File.Exists(pathFile))
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter(pathFile);
                string style = "แมโคร,มาโคร,มาโค";
                file.WriteLine(style);
                style = "เรจิสเตอร์,รีจีสเตอร์,รีจิสเตอร์";
                file.WriteLine(style);
                style = "เมทริกซ์,แมทริก,เมตทริก,";
                file.WriteLine(style);
                style = "เมท็อด,เม็ทตอด,เมทตอด";
                file.WriteLine(style);
                style = "แมนทิสซา,แมนติสซ่า,แมนทิสซ่า,แมนติทซ่า,แมนทิทซ่า";
                file.WriteLine(style);
                style = "เบราว์เซอร์,บราวเซอร์";
                file.WriteLine(style);
                style = "แอ็กทิฟ,แอ๊กทีป,เอทีฟ";
                file.WriteLine(style);
                style = "เอกซ์ทราเน็ต,แอ๊กทราเน็ต";
                file.WriteLine(style);
                style = "มอดุลาร์,มอดุลล่า";
                file.WriteLine(style);
                style = "แอนะล็อก,อนาลอก,อนาล๊อก";
                file.WriteLine(style);
                style = "เอนทิตี,เอ็นติตี้";
                file.WriteLine(style);
                style = "แอปเพล็ต,แอปเพล็ท";
                file.WriteLine(style);
                style = "มัลติ,มันติ";
                file.WriteLine(style);
                style = "อัลกอล,อันกอล";
                file.WriteLine(style);
                style = "มิดเดิล,มิเดิล";
                file.WriteLine(style);
                style = "อาร์กิวเมนต์,อะกิวเม้น";
                file.WriteLine(style);
                style = "แอสเซมเบลอร์,แอสซิมเบลอร์";
                file.WriteLine(style);
                style = "เกตเวย์,เก็ตเว";
                file.WriteLine(style);
                style = "เคส,";
                file.WriteLine(style);
                style = "โทเค็น,โทคเค่น";
                file.WriteLine(style);
                style = "เทลเน็ต,เท็วเน็ตม,เทวเน็ต";
                file.WriteLine(style);
                style = "แบนด์วิดท์,แบนวิด";
                file.WriteLine(style);
                style = "เพรดิเคต,พรีดิกเคต";
                file.WriteLine(style);
                style = "โพรโทคอล,โปรโตคอ";
                file.WriteLine(style);
                style = "อ็อบเจกต์,อ๊อปเจ๊ก";
                file.WriteLine(style);
                style = "ซ็อกเก็ต,ซ๊อกเก็ต";
                file.WriteLine(style);
                style = "ไฟร์วอลล์,ไฟล์วอ";
                file.WriteLine(style);
                style = "ดิจิทัล,ดิจิตอล";
                file.WriteLine(style);
                style = "ฟลิปฟล็อป,ฟลิ๊กฟลอป";
                file.WriteLine(style);
                style = "ลินุกซ์,ลีนุกซ์,ลีนุก";
                file.WriteLine(style);

                file.Close();
            }
            else
            {

                string[] lines = System.IO.File.ReadAllLines(pathFile);
                foreach (string line in lines)
                {
                    // Use a tab to indent each line of the file.
                    //Console.WriteLine("\t" + line);
                }
            }*/
        }

        public void CreateStyleFile(string pathFile)
        {
            if (!File.Exists(pathFile))
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter(pathFile);

                file.Write(this.style);
                file.Close();
            }
            else
            {
                string text = System.IO.File.ReadAllText(pathFile);
                if (text != this.style)
                {
                    File.Delete(pathFile);
                    CreateStyleFile(pathFile);
                }
            }
        }

        public void CreateWordFile(string pathFile)
        {
            if (!File.Exists(pathFile))
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter(pathFile);
                string[] newWords = this.words.Split('\n');
                int lenfth = newWords.Length -1;
                for (int i = 0; i < lenfth; i++)
                {
                    //style = "ลินุกซ์,ลีนุกซ์,ลีนุก";
                    file.WriteLine(newWords[i]);
                }
                file.Close();
            }
            else
            {
                string text = "";
                string[] lines = System.IO.File.ReadAllLines(pathFile);
                foreach (string line in lines)
                {
                    text += line+"\n";
                    // Use a tab to indent each line of the file.
                    //Console.WriteLine("\t" + line);
                }
                if (text != this.words)
                {
                    File.Delete(pathFile);
                    CreateWordFile(pathFile);
                }
            }
        }
    }
}
