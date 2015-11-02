using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Web.UI.HtmlControls;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Reflection;
using System.Collections;

namespace WebScriptControl
{
    [ComVisible(true)]
    public partial class WebControl : System.Web.UI.Page
    {
        ArrayList varlist = new ArrayList();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                //Load Canned Demo Script to Editor on page load
                System.Text.StringBuilder sCode = default(System.Text.StringBuilder);
                sCode = GenerateCannedRoutine();
                this.fctb.Text = sCode.ToString();
            }
        }

        //Get selected text from javascript
        StringBuilder errorJS = new StringBuilder();
        public string passJS
        {
            get
            {
                if (errorJS.Length == 0)
                {
                    return "''";
                }
                else
                    errorJS.Length -= 1;
                return "/" + errorJS + "/g";
            }
        }

        //Pass Array of keywords to  JavaScript
        private string ArrayListToString(ref ArrayList _ArrayList)
        {
            int intCount;
            string strFinal = "";
            for (intCount = 0; intCount <= _ArrayList.Count - 1; intCount++)
            {
                if (intCount > 0)
                {
                    strFinal += "~";
                }
                strFinal += _ArrayList[intCount].ToString();
            }
            return strFinal;
        }

        protected void rtxtOutput_TextChanged(object sender, EventArgs e)
        {

        }

        //Count Textbox Lines
        int SerialsCounter = 0;
        protected void fctb_TextChanged(object sender, EventArgs e)
        {
            string[] lines = Regex.Split(fctb.Text.Trim(), "\r\n");
            SerialsCounter = lines.Length;
        }

        //Canned demo script
        private StringBuilder GenerateCannedRoutine()
        {
            StringBuilder oRetSB = new StringBuilder();
            string sText = "Jan 1, 2017";
            oRetSB.Append("'Demo Script...Drag and Drop Any .txt, .vb or .vbs Here!!! " + "\r\n");
            oRetSB.Append("If Date.Parse(\"" + sText + "\") < System.DateTime.Now() Then " + "\r\n");
            oRetSB.Append(" Return \"OLD\" " + "\r\n");
            oRetSB.Append("Else " + "\r\n");
            oRetSB.Append(" Return \"NEW\" " + "\r\n");
            oRetSB.Append("End If " + "\r\n");
            return oRetSB;
        }

        public object CompileAndRunCode(string VBCodeToExecute)
        {
            string sReturn_DataType = null;
            string sReturn_Value = "";
            try
            {
                // Instance our CodeDom wrapper
                cVBEvalProvider ep = new cVBEvalProvider();

                /// Big Time
                StringBuilder sbr = new StringBuilder();
                sbr.Append("Dim Field As String = String.Empty" + "\r\n");
                sbr.Append(VBCodeToExecute + "\r\n");
                sbr.Append("Return Field");

                // Compile and run
                object objResult = ep.Eval(sbr.ToString());
                if (ep.CompilerErrors.Count != 0)
                {
                    //System.Diagnostics.Debug.WriteLine("CompileAndRunCode: Compile Error Count = " + ep.CompilerErrors.Count);
                    //System.Diagnostics.Debug.WriteLine(ep.CompilerErrors[0]);
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "error", "alert('Error: " + ep.CompilerErrors[0] + "');", true);
                    return "ERROR";
                }
                Type t = objResult.GetType();
                if (t.ToString() == "System.String")
                {
                    sReturn_DataType = t.ToString();
                    sReturn_Value = Convert.ToString(objResult);
                }
                else
                {
                    // Some other type of data - not really handled at 
                    // this point. rwd                    
                }
            }
            catch (Exception ex)
            {
                rtxtOutput.Text = sReturn_Value;
                rtxtOutput.Text = "Error: " + ex.Message;
            }
            return sReturn_Value;
        }

        protected void btnExecute_Click(object sender, EventArgs e)
        {
            try
            {
                object oRetVal = CompileAndRunCode(this.fctb.Text);
                //pass compile error to command textbox
                rtxtOutput.Text = oRetVal.ToString();
                //pass all stripped variable name from compile error to javascript to get highlighted on editor
                Regex ItemRegex = new Regex("(?<=\')[\\w]+(?!=\')", RegexOptions.Compiled);
                foreach (Match ItemMatch in ItemRegex.Matches(oRetVal.ToString()))
                {
                    if (ItemMatch.Success)
                    {
                        string stripvalue = ItemMatch.ToString();
                        varlist.Add(stripvalue);
                        txtValue.Value = ArrayListToString(ref varlist);
                        errorJS.Append(stripvalue);
                        errorJS.Append("|");
                    }
                }
                //change font color on error
                if (errorJS.Length != 0)
                {
                    rtxtOutput.ForeColor = ColorTranslator.FromHtml("#f0471c");
                }
                else
                    rtxtOutput.ForeColor = ColorTranslator.FromHtml("#000000");
            }
            catch (Exception x)
            {
                rtxtOutput.ForeColor = ColorTranslator.FromHtml("#f0471c");
                rtxtOutput.Text = x.Message;
            }
        }

        protected void fctb_Load(object sender, EventArgs e)
        {

        }

        protected void btnClear_Click(object sender, EventArgs e)
        {
            fctb.Text = string.Empty;
            rtxtOutput.Text = string.Empty;
        }

        protected void rtxtOutput_Load(object sender, EventArgs e)
        {

        }

        //Button to Declare Variables and Add Empty Lines on Top
        protected void btnDeclare_Click(object sender, EventArgs e)
        {
            string declareVars = "'Declare and Initialize Variables Here!";
            StringBuilder varsDecl = new StringBuilder();
            varsDecl.Append(declareVars + "\n");
            varsDecl.Append("\n\n\n\n\n");
            varsDecl.Append(fctb.Text);
            bool exist = fctb.Text.Contains(declareVars);
            if (!exist)
            {
                fctb.Text = varsDecl.ToString();
            }
        }

        //Button to Save the Script locally
        protected void buttonSave_Click(object sender, EventArgs e)
        {
            string fileName = hidValue.Value;

            if (fileName != string.Empty)
            {
                Response.Clear();
                Response.ContentType = "text/plain";
                Response.AppendHeader("Content-Disposition", "attachment; filename=" + fileName);
                Response.Write(fctb.Text);
                Response.End();
            }
            else
            {
                Response.Clear();
                Response.ContentType = "text/plain";
                Response.AppendHeader("Content-Disposition", "attachment; filename=" + "my_script_" + DateTime.Now + ".vb");
                Response.Write(fctb.Text);
                Response.End();
            }
        }
    }
}
