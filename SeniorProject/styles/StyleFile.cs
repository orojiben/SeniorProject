using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace SeniorProject
{
    static class StyleFile
    {
        //private static string path = @"styles\word_style.xml";
        //private static string path = @"C:\styles\word_style.xml";
        private static string folder = "CheckingThesis";
        private static string pathDoc = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
        private static string path = pathDoc + "\\" + folder + "\\stlyes.xml";


        static public void CheckCreateFile()
        {
            if (!File.Exists(path))
            {
                XmlDocument xmlDoc = new XmlDocument();
                XmlNode rootNode = xmlDoc.CreateElement("styles");
                xmlDoc.AppendChild(rootNode);
                xmlDoc.Save(path);
            }
        }

        public static XmlDocument LoadStylePath()
        {
            XmlDocument xmlDoc = new XmlDocument();
            try
            {
                xmlDoc.Load(path);
       
            }
            catch
            {
                return LoadStylePath();
            }
            return xmlDoc;
        }

        public static List<Styles> LoadStyle()
        {
            CreateStyle cs = new CreateStyle();
            cs.Check();
            List<Styles> styles = new List<Styles>();
            XmlDocument xmlDoc = LoadStylePath();
            //xmlDoc.Load(path);
            XmlNodeList userNodes = xmlDoc.SelectNodes("//styles");

            XmlNode rootNodes = userNodes[0];
            foreach (XmlNode rootNode in rootNodes)
            {
                Styles style = new Styles();
                foreach (XmlNode nodeClass1 in rootNode.ChildNodes)
                {

                   /*if (nodeClass1.LocalName == "dictionarys")
                    {
                        foreach (XmlNode nodeClass2 in nodeClass1.ChildNodes)
                        {
                            // Console.Write("\n" + nodeClass2.InnerText + "\n");
                            style.addDictionary(nodeClass2.InnerText);
                        }
                        styles.Add(style);
                    }
                    else  */if (nodeClass1.LocalName == "fonts")
                    {
                        foreach (XmlNode nodeClass2 in nodeClass1.ChildNodes)
                        {
                            // Console.Write("\n" + nodeClass2.InnerText + "\n");
                            //style.addFont(nodeClass2.InnerText);
                            string fontName = "";
                            foreach (XmlNode nodeClass3 in nodeClass2.ChildNodes)
                            {
                                if (fontName == "")
                                {
                                    fontName = nodeClass3.InnerText;
                                }
                                else
                                {
                                    float coverTitle = 0.0f;
                                    float coverOperator = 0.0f;
                                    float chapter = 0.0f;
                                    float namechapter = 0.0f;
                                    float topics = 0.0f;
                                    float subheading = 0.0f;
                                    float substance = 0.0f;
                                    string fontNameLanguage  = nodeClass3.LocalName;
                                    foreach (XmlNode nodeClass4 in nodeClass3.ChildNodes)
                                    {
                                        if (nodeClass4.LocalName == "CoverTitle")
                                        {
                                            coverTitle = float.Parse(nodeClass4.InnerText);
                                        }
                                        else if (nodeClass4.LocalName == "CoverOperator")
                                        {
                                            coverOperator = float.Parse(nodeClass4.InnerText);
                                        }
                                        else if (nodeClass4.LocalName == "Chapter")
                                        {
                                            chapter = float.Parse(nodeClass4.InnerText);
                                        }
                                        else if (nodeClass4.LocalName == "Namechapter")
                                        {
                                            namechapter = float.Parse(nodeClass4.InnerText);
                                        }
                                        else if (nodeClass4.LocalName == "Topics")
                                        {
                                            topics = float.Parse(nodeClass4.InnerText);
                                        }
                                        else if (nodeClass4.LocalName == "Subheading")
                                        {
                                            subheading = float.Parse(nodeClass4.InnerText);
                                        }
                                        else if (nodeClass4.LocalName == "Substance")
                                        {
                                            substance = float.Parse(nodeClass4.InnerText);
                                            style.addFont( fontName,  fontNameLanguage,  coverTitle,  coverOperator,  chapter,
             namechapter,  topics,  subheading,  substance);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else if (nodeClass1.LocalName == "Departments")
                    {
                        foreach (XmlNode nodeClass2 in nodeClass1.ChildNodes)
                        {
                            // Console.Write("\n" + nodeClass2.InnerText + "\n");
                            style.addDepartment(nodeClass2.InnerText);
                        }
                    }
                    else if (nodeClass1.LocalName == "Margin")
                    {
                        style.Margin = nodeClass1.InnerText;
                    }
                    else if (nodeClass1.LocalName == "Paper")
                    {
                        style.Paper = nodeClass1.InnerText;
                    }
                    else if (nodeClass1.LocalName == "Indent")
                    {
                        style.Indent =float.Parse(nodeClass1.InnerText);
                        styles.Add(style);
                    }
                    else
                    {
                        style.Name = nodeClass1.InnerText;
                        // Console.Write("\n" + nodeClass1.InnerText + "\n");
                    }

                }
            }

            return styles;
        }



        public static void WriteStyle(string title)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(path);
            XmlNodeList userNodes = xmlDoc.SelectNodes("//styles");

            XmlNode rootNode = userNodes[0];
            XmlNode styleNode = xmlDoc.CreateElement("style");
            XmlNode nameNode = xmlDoc.CreateElement("Nameformat");
            XmlNode marginNode = xmlDoc.CreateElement("Margin");
            XmlNode fontsNode = xmlDoc.CreateElement("fonts");
            XmlNode fontNode = xmlDoc.CreateElement("font");

            fontNode.InnerText = "new";
            fontsNode.AppendChild(fontNode);

            XmlNode dictionarysNode = xmlDoc.CreateElement("dictionarys");
            XmlNode dictionaryNode = xmlDoc.CreateElement("dictionary");
            dictionaryNode.InnerText = "thai";
            dictionarysNode.AppendChild(dictionaryNode);
            // XmlAttribute attribute = xmlDoc.CreateAttribute("Nameformat");
            // attribute.Value = "Engineering";
            //  user.Attributes.Append(attribute);
            nameNode.InnerText = "Engineering";
            marginNode.InnerText = "2,2,2,2";
            styleNode.AppendChild(nameNode);
            styleNode.AppendChild(marginNode);
            styleNode.AppendChild(fontsNode);
            styleNode.AppendChild(dictionarysNode);
            rootNode.AppendChild(styleNode);

            xmlDoc.Save(path);
        }

        public static void WriteStyle(Styles style)
        {
            if (style != null)
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(path);
                XmlNodeList userNodes = xmlDoc.SelectNodes("//styles");

                XmlNode rootNode = userNodes[0];
                XmlNode styleNode = xmlDoc.CreateElement("style");
                XmlNode nameNode = xmlDoc.CreateElement("Nameformat");
                XmlNode marginNode = xmlDoc.CreateElement("Margin");
                XmlNode fontsNode = xmlDoc.CreateElement("fonts");

                foreach (StyleFont font in style.StyleFont)
                {
                    XmlNode fontNode = xmlDoc.CreateElement("font");
                    //fontNode.InnerText = font;
                    fontsNode.AppendChild(fontNode);
                }

                XmlNode dictionarysNode = xmlDoc.CreateElement("dictionarys");
               /* foreach (string dictionary in style.Dictionarys)
                {
                    XmlNode dictionaryNode = xmlDoc.CreateElement("dictionary");
                    dictionaryNode.InnerText = dictionary;
                    dictionarysNode.AppendChild(dictionaryNode);
                }*/
                // XmlAttribute attribute = xmlDoc.CreateAttribute("Nameformat");
                // attribute.Value = "Engineering";
                //  user.Attributes.Append(attribute);
                nameNode.InnerText = style.Name;
                marginNode.InnerText = style.Margin;
                styleNode.AppendChild(nameNode);
                styleNode.AppendChild(marginNode);
                styleNode.AppendChild(fontsNode);
                styleNode.AppendChild(dictionarysNode);
                rootNode.AppendChild(styleNode);

                xmlDoc.Save(path);
            }
        }


        public static void EditStyle(Styles style, string nameOld)
        {
            if (style != null)
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(path);
                XmlNodeList userNodes = xmlDoc.SelectNodes("//styles");

                XmlNode rootNode = userNodes[0];

                foreach (XmlNode node in rootNode.ChildNodes)
                {
                    if (node.ChildNodes[0].InnerText == nameOld)
                    {
                        node.ChildNodes[0].InnerText = style.Name;
                        node.ChildNodes[1].InnerText = style.Margin;
                        node.ChildNodes[2].RemoveAll();
                        XmlNode fontsNode = xmlDoc.CreateElement("fonts");
                        foreach (StyleFont value in style.StyleFont)
                        {
                            XmlNode fontNode = xmlDoc.CreateElement("font");
                            //fontNode.InnerText = value;
                            node.ChildNodes[2].AppendChild(fontNode);
                        }
                        node.ChildNodes[3].RemoveAll();
                        /*foreach (string value in style.Dictionarys)
                        {
                            XmlNode dictionaryNode = xmlDoc.CreateElement("dictionary");
                            dictionaryNode.InnerText = value;
                            node.ChildNodes[3].AppendChild(dictionaryNode);
                        }*/
                        break;
                    }
                }


                xmlDoc.Save(path);
            }
        }


        public static void RemoveStyle(string nameOld)
        {

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(path);
            XmlNodeList userNodes = xmlDoc.SelectNodes("//styles");

            XmlNode rootNode = userNodes[0];

            foreach (XmlNode node in rootNode.ChildNodes)
            {
                if (node.ChildNodes[0].InnerText == nameOld)
                {
                    rootNode.RemoveChild(node);
                    break;
                }

            }


            xmlDoc.Save(path);

        }

        public static void WriteXML(string title)
        {

            using (XmlWriter writer = XmlWriter.Create(@"..\..\style\text.xml"))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Employees");
                writer.WriteStartElement("Employee");

                writer.WriteElementString("ID", "006");
                writer.WriteElementString("FirstName", "oroji");
                writer.WriteElementString("LastName", "ben");
                writer.WriteElementString("Salary", "Not");

                writer.WriteEndElement();

                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        public static void WriteXML2(string title)
        {

            XmlDocument xmlDoc = new XmlDocument();
            XmlNode rootNode = xmlDoc.CreateElement("users");
            xmlDoc.AppendChild(rootNode);

            XmlNode userNode = xmlDoc.CreateElement("user");
            XmlAttribute attribute = xmlDoc.CreateAttribute("age");
            attribute.Value = "42";
            userNode.Attributes.Append(attribute);
            userNode.InnerText = "John Doe";
            rootNode.AppendChild(userNode);

            userNode = xmlDoc.CreateElement("user");
            attribute = xmlDoc.CreateAttribute("age");
            attribute.Value = "39";
            userNode.Attributes.Append(attribute);
            userNode.InnerText = "Jane Doe";
            rootNode.AppendChild(userNode);

            xmlDoc.Save(@"..\..\style\test-doc.xml");
        }

        public static void WriteXML3(string title)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(@"..\..\style\test-doc.xml");
            XmlNodeList userNodes = xmlDoc.SelectNodes("//users/user");
            foreach (XmlNode userNode in userNodes)
            {
                int age = int.Parse(userNode.Attributes["age"].Value);
                userNode.Attributes["age"].Value = (age + 1).ToString();
                // userNode.InnerText = "Orojiben";
            }
            XmlNodeList userAdd = xmlDoc.SelectNodes("//users");
            XmlNode forAdd = userAdd[0];
            XmlNode user = xmlDoc.CreateElement("user");
            XmlAttribute attribute = xmlDoc.CreateAttribute("age");
            attribute.Value = "42";
            user.Attributes.Append(attribute);
            user.InnerText = "Ben Doe";
            forAdd.AppendChild(user);

            xmlDoc.Save(@"..\..\style\test-doc.xml");
        }
    }
}
