using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.CodeDom.Compiler;
using System.Reflection;
using System.Linq;
using System.Web.Configuration;

namespace WebScriptControl
{
    public class cVBEvalProvider
    {
        private CompilerErrorCollection m_oCompilerErrors;
        public CompilerErrorCollection CompilerErrors
        {
            get
            {
                return m_oCompilerErrors;
            }

            set
            {
                m_oCompilerErrors = value;
            }
        }

        public cVBEvalProvider() : base()
        {
            m_oCompilerErrors = new CompilerErrorCollection();
        }

        public object Eval(string vbCode)
        {
            VBCodeProvider oCodeProvider = new VBCodeProvider();
            // Obsolete in 2.0 framework
            // Dim oICCompiler As ICodeCompiler = oCodeProvider.CreateCompiler
            CompilerParameters oCParams = new CompilerParameters();
            CompilerResults oCResults = default(CompilerResults);
            System.Reflection.Assembly oAssy = default(System.Reflection.Assembly);
            object oExecInstance = null;
            object oRetObj = null;
            MethodInfo oMethodInfo = default(MethodInfo);
            Type oType = default(Type);
            List<string> strb = new List<string>();
            string dlls = WebConfigurationManager.AppSettings["dllpath"].ToString();
            try
            {
                // Setup the Compiler Parameters  
                // Add any referenced assemblies
                oCParams.ReferencedAssemblies.Add("system.dll");
                oCParams.ReferencedAssemblies.Add("system.xml.dll");
                oCParams.ReferencedAssemblies.Add("system.data.dll");
                oCParams.ReferencedAssemblies.Add("system.io.dll");
                oCParams.ReferencedAssemblies.Add("system.web.dll");
                oCParams.ReferencedAssemblies.Add("system.collections.dll");
                oCParams.ReferencedAssemblies.Add("Microsoft.VisualBasic.dll");
                oCParams.ReferencedAssemblies.Add("System.Data.DataSetExtensions.dll");
                oCParams.ReferencedAssemblies.Add("System.Core.dll");
                oCParams.ReferencedAssemblies.Add("System.Data.Linq.dll");
                oCParams.ReferencedAssemblies.Add("System.Windows.Forms.dll");
                oCParams.ReferencedAssemblies.Add(dlls);
                oCParams.CompilerOptions = "/t:library /optimize";
                oCParams.GenerateInMemory = true;
                oCParams.GenerateExecutable = false;
                oCParams.IncludeDebugInformation = true;
                // Generate the Code Framework
                StringBuilder sb = new StringBuilder("");
                sb.Append("Imports System" + "\r\n");
                sb.Append("Imports Parse" + "\r\n");
                sb.Append("Imports System.Xml" + "\r\n");
                sb.Append("Imports System.Data" + "\r\n");
                sb.Append("Imports System.IO" + "\r\n");
                sb.Append("Imports Microsoft.VisualBasic" + "\r\n");
                sb.Append("Imports Microsoft.VisualBasic.DateAndTime" + "\r\n");
                sb.Append("Imports System.Collections.Generic" + "\r\n");
                sb.Append("Imports System.Data.Linq" + "\r\n");
                sb.Append("Imports System.Windows.Forms" + "\r\n");
                // Build a little wrapper code, with our passed in code in the middle 
                sb.Append("Namespace dValuate" + "\r\n");
                sb.Append("Class EvalRunTime " + "\r\n");
                sb.Append("Public Function EvaluateIt() As Object " + "\r\n");
                //sb.Append(GetList());
                sb.Append(vbCode + "\r\n");
                sb.Append("End Function " + "\r\n");
                sb.Append("End Class " + "\r\n");
                sb.Append("End Namespace" + "\r\n");
                //Debug.WriteLine(sb.ToString());
                try
                {
                    // Compile and get results 
                    // 2.0 Framework - Method called from Code Provider
                    oCResults = oCodeProvider.CompileAssemblyFromSource(oCParams, sb.ToString());
                    // 1.1 Framework - Method called from CodeCompiler Interface
                    // cr = oICCompiler.CompileAssemblyFromSource (cp, sb.ToString)
                    // Check for compile time errors
                    if (oCResults.Errors.Count > 0)
                    {
                        for (int i = 0; i < oCResults.Errors.Count; i++)
                        {
                            strb.Add(oCResults.Errors[i].ToString().Substring(oCResults.Errors[i].ToString().IndexOf("error ")) + Environment.NewLine);
                        }
                        //return oCResults;
                    }
                    else
                    {
                        // No Errors On Compile, so continue to process...
                        oAssy = oCResults.CompiledAssembly;
                        oExecInstance = oAssy.CreateInstance("dValuate.EvalRunTime");
                        oType = oExecInstance.GetType();
                        oMethodInfo = oType.GetMethod("EvaluateIt");
                        oRetObj = oMethodInfo.Invoke(oExecInstance, null);
                        return oRetObj;
                    }
                }
                catch (Exception ex)
                {
                    // Runtime Errors Are Caught Here
                    strb.Add("Error Runtime: " + ex.InnerException.Message.ToString() + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                strb.Add(ex.InnerException.Message.ToString() + Environment.NewLine);
            }

            //return oRetObj;
            StringBuilder result = new StringBuilder();
            var distinct = (
                from item in strb
                orderby item
                select item).Distinct();
            foreach (string value in distinct)
            {
                result.Append(value.ToString());
            }
            return result.ToString();
        }
    }
}
