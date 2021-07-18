using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.IO;

using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Windows.Forms;


namespace ToDoNotificheBSC
{
    public class SignFiles
    {
        public static string fileRTF = "";
        public static string filePNG = "";
        public static bool testSignPDF = false;

        public static int tipofirma = -1;

        public static string tempOriginalPDF = Path.Combine(Path.GetTempPath(), @"t.pdf");
        public static string tempSignedPDF = Path.Combine(Path.GetTempPath(), @"newt.pdf");
        public static string tempMergePDF = Path.Combine(Path.GetTempPath(), @"tempMerge.pdf");

        public static string startXML = Path.Combine(System.Windows.Forms.Application.StartupPath, "test.knos-fr");
        public static string startXML_baseurl = "";
        public static int startXML_idobject = 0;
        public static int startXML_idobject_certificato = 0;


        // stati
        public static int KnoS_Certificato_IdStatusIniziale = 0;
        public static int KnoS_Certificato_IdStatus1F = 0;
        public static int KnoS_Certificato_IdStatus2F = 0;
        public static int KnoS_PDL_IdStatusPDLFirmato = 0;
        public static int KnoS_PDL_IdStatusPDLDaFirmare = 0;

        public static int KnoS_IdActionPDFdaFirmare = 0;
        public static int KnoS_IdActionPDFFirmato = 0;

    

        public static string GetStringFromPNG(string path, int width, int height)
        { 
            MemoryStream stream = new MemoryStream();
            string newPath = Path.Combine(Environment.CurrentDirectory, path);
            Image img = Image.FromFile(newPath);
            img.Save(stream, System.Drawing.Imaging.ImageFormat.Png);

            byte [] bytes = stream.ToArray();

            string str = BitConverter.ToString(bytes, 0).Replace("-", string.Empty);
            //string str = System.Text.Encoding.UTF8.GetString(bytes);

        //    string mpic = @"{\pict\pngblip\picw" + 
        //        img.Width.ToString() + @"\pich" + img.Height.ToString() +
        //        @"\picwgoa" + width.ToString() + @"\pichgoa" + height.ToString() + 
         //       @"\hex " + str + "}";
	        string mpic = @"{\pict\pngblip\picw" + 
		        img.Width.ToString() + @"\pich" + img.Height.ToString() +
		        @"\picwgoal" + width.ToString() + @"\pichgoal" + height.ToString() + 
		        @"\bin " + str + "}";

            return mpic;        
        }



        public static bool LoadInfoXML(string fileInfoXML)
        {
            try
            {
                // SECTION 1. Create a DOM Document and load the XML data into it.
                XmlDocument dom = new XmlDocument();
                dom.Load(fileInfoXML);

                XmlNode nInfo = dom.SelectSingleNode("KnosEnvelope");
                startXML_baseurl = nInfo.Attributes["knosBaseUrl"].Value;

                nInfo = dom.SelectSingleNode("KnosEnvelope/PDL");                
                int.TryParse(nInfo.Attributes["IdObject"].Value.ToString(), out startXML_idobject);
                
                try
                {
                    if (nInfo.Attributes.GetNamedItem("IdObjectCertificato") != null)
                    {
                        int.TryParse(nInfo.Attributes["IdObjectCertificato"].Value.ToString(), out startXML_idobject_certificato);
                    }
                }
                catch (XmlException xmlEx)
                {
                }

                return true;
            }
            catch (XmlException xmlEx)
            {
                MessageBox.Show(xmlEx.Message);
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Impossibile leggere il file con le informazioni dei lotti da gestire!");
                return false;
            }

        }


        /*
        private void AddNodeLotto(XmlNode inXmlNode, TreeNode inTreeNode)
        {
            XmlNode xNode;
            TreeNode tNode;
            XmlNodeList nodeList;

            int i;
            string attributeList;
            string nodeName;
            object nodeTag;

            // Loop through the XML nodes until the leaf is reached.
            // Add the nodes to the TreeView during the looping process.
            if (inXmlNode.HasChildNodes)
            {


                nodeList = inXmlNode.ChildNodes;
                for (i = 0; i <= nodeList.Count - 1; i++)
                {
                    xNode = inXmlNode.ChildNodes[i];

                    attributeList = "";
                    nodeName = "";
                    nodeTag = null;

                    nodeName = xNode.Name;

                    switch (xNode.Name)
                    {
                        case "LinkLottoCS":

                            // per ciascun lotto creo le directory di appoggio nel caso manchino
                            LoadFolderLotto(Path.Combine(xNode.Attributes["netPathCS"].Value, xNode.Attributes["pathLottoCS"].Value));

                            nodeName = string.Format("{0} - {1}", xNode.Attributes["name"].Value, xNode.Attributes["description"].Value);
                            nodeTag = new CustomNodeData(xNode.Name, Path.Combine(xNode.Attributes["netPathCS"].Value, xNode.Attributes["pathLottoCS"].Value), xNode.InnerXml, xNode.OuterXml, 1);



                            try
                            {
                                xmlDocLotto.Load(Path.Combine(Path.Combine(xNode.Attributes["netPathCS"].Value, xNode.Attributes["pathLottoCS"].Value), "originali", "info.xml"));

                                inTreeNode.Nodes.Add(new TreeNode(nodeName));
                                tNode = inTreeNode.Nodes[i];
                                tNode.Tag = nodeTag;
                                tNode.ImageIndex = 1;
                                tNode.SelectedImageIndex = 2;




                                XmlNode nInfoLotto = xmlDocLotto.SelectSingleNode("KnosEnvelope/LottoCS");






                                //inTreeNode.ImageIndex = 1;
                                if (nInfoLotto.Attributes["CodiceMarca"] != null)
                                {
                                    tNode.ImageIndex = 3;

                                }

                                lottoIdObject = int.Parse(nInfoLotto.Attributes["idObject"].Value.ToString());
                                lottoTokenCS = int.Parse(nInfoLotto.Attributes["tokenCS"].Value.ToString());

                                AddNode(xNode, tNode);

                                AddNode(xmlDocLotto.SelectSingleNode("KnosEnvelope/LottoCS"), tNode);

                                //textBox1.Text = xmlDocLotto.OuterXml.ToString();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }


                            break;

                        default:

                            break;

                    }
                }
            }
            else
            {
                // Here you need to pull the data from the XmlNode based on the
                // type of node, whether attribute values are required, and so forth.

                //inTreeNode.Text = (inXmlNode.OuterXml).Trim();


            }
        }


        private void AddNode(XmlNode inXmlNode, TreeNode inTreeNode)
        {
            XmlNode xNode;
            TreeNode tNode;
            XmlNodeList nodeList;

            int imgListIndex = 0;

            int i;
            string attributeList;
            string nodeName;
            object nodeTag;

            Application.DoEvents();

            ResumeLayout();

            try
            {

                // Loop through the XML nodes until the leaf is reached.
                // Add the nodes to the TreeView during the looping process.
                if (inXmlNode.HasChildNodes)
                {


                    nodeList = inXmlNode.ChildNodes;
                    for (i = 0; i <= nodeList.Count - 1; i++)
                    {
                        xNode = inXmlNode.ChildNodes[i];

                        attributeList = "";
                        nodeName = "";
                        nodeTag = null;

                        nodeName = xNode.Name;

                        switch (xNode.Name)
                        {
                            case "LinkLottoCS":

                                // per ciascun lotto creo le directory di appoggio nel caso manchino
                                LoadFolderLotto(Path.Combine(xNode.Attributes["netPathCS"].Value, xNode.Attributes["pathLottoCS"].Value));

                                imgListIndex = 3;
                                nodeName = string.Format("{0} - {1}", xNode.Attributes["name"].Value, xNode.Attributes["description"].Value);
                                nodeTag = new CustomNodeData(xNode.Name, Path.Combine(xNode.Attributes["netPathCS"].Value, xNode.Attributes["pathLottoCS"].Value), xNode.InnerXml, xNode.OuterXml, 1);

                                if (xNode.Attributes["DocSignedHash"].Value.ToString() != null)
                                {

                                }

                                break;

                            case "Pubblicazioni":

                                nodeName = "Elenco Pubblicazioni";
                                imgListIndex = 4;
                                //inTreeNode.ImageIndex = 4;

                                nodeTag = new CustomNodeData(xNode.Name, "", xNode.InnerXml, xNode.OuterXml, 1);

                                break;

                            case "Pubblicazione":

                                nodeName = "Pubblicazione";

                                foreach (XmlAttribute attr in xNode.Attributes)
                                {
                                    attributeList += string.Format(" | [{0} : {1}]", attr.Name, attr.Value);
                                    nodeName = string.Format("{0} - {1}", xNode.Name, attributeList);
                                }

                                imgListIndex = 5;
                                nodeTag = new CustomNodeData(xNode.Name, "", xNode.InnerXml, xNode.OuterXml, 1);

                                break;

                            case "Attributi":

                                imgListIndex = 6;

                                nodeName = "Attributi";
                                nodeTag = new CustomNodeData(xNode.Name, "", xNode.InnerXml, xNode.OuterXml, 1);

                                foreach (XmlNode xaNode in xNode.ChildNodes)
                                {

                                    attributeList = string.Format(" {0} : {1}", xaNode.Attributes["Nome"].Value.ToString(), xaNode.Attributes["Valore"].Value.ToString());
                                    inTreeNode.Nodes.Add(attributeList, attributeList, imgListIndex);

                                    //inTreeNode.ImageIndex = 5;
                                    //inTreeNode.SelectedImageIndex = 5;                           
                                }


                                break;

                            case "Attributo":

                                break;

                            case "Documenti":

                                nodeName = "Documento";

                                nodeTag = new CustomNodeData(xNode.Name, "", xNode.InnerXml, xNode.OuterXml, 1);



                                foreach (XmlNode xaNode in xNode.ChildNodes)
                                {

                                    imgListIndex = 7;



                                    if (xaNode.Attributes["DocSignedHash"].Value.ToString() != "")
                                    {
                                        attributeList = string.Format(" {0} : {1}", xaNode.Attributes["FileName"].Value.ToString(), "Firmato");
                                        imgListIndex += 1;
                                    }
                                    else
                                    {
                                        attributeList = string.Format(" {0}", xaNode.Attributes["FileName"].Value.ToString());
                                    }


                                    inTreeNode.Nodes.Add(attributeList, attributeList, imgListIndex);

                                }


                                break;


                            case "Documento":

                                nodeName = "Documento";

                                nodeTag = new CustomNodeData(xNode.Name, "", xNode.InnerXml, xNode.OuterXml, 1);

                                break;

                            default:

                                foreach (XmlAttribute attr in xNode.Attributes)
                                {
                                    attributeList += string.Format(" | [{0} : {1}]", attr.Name, attr.Value);
                                    nodeName = string.Format("{0} - {1}", xNode.Name, attributeList);
                                }

                                break;

                        }

                        if ((xNode.Name != "Attributi") && (xNode.Name != "Documenti"))
                        {
                            inTreeNode.Nodes.Add(new TreeNode(nodeName));
                            tNode = inTreeNode.Nodes[i];
                            tNode.ImageIndex = imgListIndex;
                            tNode.SelectedImageIndex = imgListIndex;
                            tNode.Tag = nodeTag;
                            AddNode(xNode, tNode);
                        }
                    }
                }
                else
                {
                    // Here you need to pull the data from the XmlNode based on the
                    // type of node, whether attribute values are required, and so forth.
                    if (inXmlNode.Attributes["Name"] != null)
                    {
                        nodeName = string.Format("{0} - {1}", inXmlNode.Attributes["Name"].Value, inXmlNode.Attributes["Value"].Value);
                        inTreeNode.Text = nodeName;// (inXmlNode.OuterXml).Trim();
                        inTreeNode.SelectedImageIndex = 2;
                    }
                    //inTreeNode.ImageIndex = 1;



                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(inXmlNode.OuterXml + " \r\n" + ex.Message);

            }
        }
        */


    }



}
