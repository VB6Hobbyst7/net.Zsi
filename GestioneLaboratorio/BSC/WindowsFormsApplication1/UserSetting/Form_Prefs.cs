using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.Design;
using System.IO;
using System.Reflection;
using System.Diagnostics;

namespace SignRTFPDF
{
    public partial class Form_Prefs : Form
    {
        public Form_Prefs()
        {
            InitializeComponent();
        }

        private void Form_Prefs_Load(object sender, EventArgs e)
        {
            //01/20/09 build table of properties using current values from Settings object
            //         all this is required because PropertyGrid doesn't allow arbitrary string property names by default.
            //         Uses the PropertySpec & PropertyTable classes from PropertyBag.cs, courtesy of Tony Allowatt 
            //         the basic constructor signature is:
            //         public PropertySpec(string propname, string name, string type, string category, 
            //                  string description, object defaultValue, Type editor, Type typeConverter)
            //         'propname' values come from Settings.Designer.cs, which is managed automatically by the project settings dialog
 
            //create & fill the table.  
            PropertyTable proptable = new PropertyTable();

            //Construct PropertyTable entries from Settings class user-scoped properties 
            SignRTFPDF.Properties.Settings settings = SignRTFPDF.Properties.Settings.Default;
            Type type = typeof(Properties.Settings);
            MemberInfo[] pi = type.GetProperties();
            foreach (MemberInfo m in pi)
            {
                
                Object[] myAttributes = m.GetCustomAttributes(true);
                if (myAttributes.Length > 0)
                {
                    for (int j = 0; j < myAttributes.Length; j++)
                    {
                        if( myAttributes[j].ToString() == "System.Configuration.UserScopedSettingAttribute")
                        {
                            PropertySpec ps = new PropertySpec("property name", "System.String");
                            switch (m.Name)
                            {
                                case "startXML":
                                    ps = new PropertySpec(
                                    "File processato",
                                    "System.Reflection.RuntimePropertyInfo",
                                    "File Locations",
                                    "File processato",
                                    settings.startXML,
                                    typeof(System.Windows.Forms.Design.FileNameEditor),
                                    typeof(System.Convert));
                                    break;

                                case "KnoS_NomeStatoIniziale":
                                    ps = new PropertySpec(
                                    "proprietà immagine segnaposto firma del capo commessa",
                                    "System.Reflection.RuntimePropertyInfo",
                                    "KnoS Settings",
                                    "Stato WF iniziale certificato",
                                    settings.KnoS_IdStatoIniziale,
                                    typeof(System.Windows.Forms.TextBox),
                                    typeof(System.Convert));
                                    break;

                                case "KnoS_NomeStato1F":
                                    ps = new PropertySpec(
                                    "Stato WF prima firma del certificato",
                                    "System.Reflection.RuntimePropertyInfo",
                                    "KnoS Settings",
                                    "Stato WF prima firma del certificato",
                                    settings.KnoS_IdStato1F,
                                    typeof(System.String),
                                    typeof(System.Convert));
                                    break;

                                case "KnoS_NomeStato2F":
                                    ps = new PropertySpec(
                                    "Stato WF seconda firma del certificato",
                                    "System.Reflection.RuntimePropertyInfo",
                                    "KnoS Settings",
                                    "Stato WF seconda firma del certificato",
                                    settings.KnoS_IdStato2F,
                                    typeof(System.String),
                                    typeof(System.Convert));
                                    break;

                                case "KnoS_NomeStatoPDFFirmato":
                                    ps = new PropertySpec(
                                    "Stato WF PDF FIRMATO",
                                    "System.Reflection.RuntimePropertyInfo",
                                    "KnoS Settings",
                                    "Stato WF PDF FIRMATO",
                                    settings.KnoS_IdStatoPDLFirmato,
                                    typeof(System.String),
                                    typeof(System.Convert));
                                    break;


                                /* 
//Files category
case "firmacapocommessa":
ps = new PropertySpec( 
"proprietà immagine segnaposto firma del capo commessa",
"System.Reflection.RuntimePropertyInfo",
"File Locations",
"NEC-BSC CEM Code Location", 
settings.firmacapocommessa,
typeof(System.Windows.Forms.Design.FileNameEditor), 
typeof(System.Convert));
break;

case "firmaresptecnico":
ps = new PropertySpec(
"Bsc CEM Code",
"System.String",
"File Locations",
"NEC-BSC CEM Code Location",
settings.firmaresptecnico.ToString(),
typeof(System.Windows.Forms.Design.FileNameEditor),
typeof(System.Convert));
break;

case "firmatecnico":
ps = new PropertySpec(
"Bsc CEM Code",
"System.String",
"File Locations",
"NEC-BSC CEM Code Location",
settings.firmatecnico.ToString(),
typeof(System.Windows.Forms.Design.FileNameEditor),
typeof(System.Convert));
break;

//Colors
                               
* case "pec_color":
ps = new PropertySpec(
"PEC Color",
typeof(System.Drawing.Color),
"Colors",
"Color used for PEC model elements",
settings.pec_color);
break;

//Fonts
case "default_plot_font":
ps = new PropertySpec(
"Default Plot Font",
typeof(Font),
"Fonts",
"Default font used for 2-D plots",
settings.default_plot_font);
break;*/

                                default:
                                    ps = new PropertySpec(m.Name, 
                                        typeof(System.String),
                                        m.Name,
                                        m.ToString(),
                                        "");
                                    break;
                                    
                            }
                            proptable.Properties.Add(ps);
                        }
                    }
                }
            }

            //this line binds the PropertyTable object to the preferences PropertyGrid control
            this.pg_Prefs.SelectedObject = proptable;
        }

        private void btn_OK_Click(object sender, EventArgs e)
        {
            //write property values back to Settings object properties

            Button btn = (Button)sender;
            Form_Prefs form = (Form_Prefs)btn.Parent;
            PropertyGrid pg = form.pg_Prefs;

            PropertyTable proptable = pg.SelectedObject as PropertyTable;
            //EMWorkbench.Properties.Settings settings = EMWorkbench.Properties.Settings.Default;

            //get the grid root
            GridItem gi = pg.SelectedGridItem;
            while (gi.Parent != null)
            {
                gi = gi.Parent;
            }

            //transfer all grid item values to Settings class properties
            foreach( GridItem item in gi.GridItems)
            {
                ParseGridItems(item); //recursive
            }

            this.Close();
        }

        private void pg_Prefs_PropertyValueChanged(object s, PropertyValueChangedEventArgs e)
        {
            Trace.WriteLine(e.ChangedItem.Label);
        }

        private void ParseGridItems(GridItem gi)
        {
            SignRTFPDF.Properties.Settings settings = SignRTFPDF.Properties.Settings.Default;

            if (gi.GridItemType == GridItemType.Category)
            {
                foreach (GridItem item in gi.GridItems)
                {
                    ParseGridItems(item); //terminates at 1st Property
                }
            }

            switch (gi.Label)
            {
                    /*
                case "firmatecnico":
                    settings.firmatecnico = gi.Value.ToString();
                    break;
                case "firmaresponsabiletecnico":
                    settings.firmaresponsabiletecnico = (Color)gi.Value;
                    break;
                case "firmacapocommessa":
                    settings.firmacapocommessa = (Font)gi.Value;
                    break;
                     */

                default:
                    break;
            }
       }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            this.Close(); //just close w/o doing anything
        }
    }
}