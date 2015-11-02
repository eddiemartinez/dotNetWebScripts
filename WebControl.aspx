<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebControl.aspx.cs" Inherits="WebScriptControl.WebControl" %>

<!DOCTYPE html>
<head id="Head1" runat="server">
    <title>ASP NET VB Scripts ACE Editor</title>
    <link rel="shortcut icon" href="images/vs.ico" />
    <script src="/ace-builds/src-noconflict/ace.js" type="text/javascript" charset="utf-8"></script>
    <script src="Scripts/jquery-1.7.2.min.js" type="text/javascript" charset="utf-8"></script>
    <script src="Scripts/modernizr-2.5.3.js" type="text/javascript" charset="utf-8"></script>
    <link href="Styles/MyStyle.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <div class="panel">
        <div id="editor">
        </div>
        <form id="form1" runat="server">
            <asp:TextBox ID="fctb" name="fctb" runat="server" hidden="true" OnTextChanged="fctb_TextChanged"
                AutoPostBack="true" TextMode="MultiLine" Height="0px" Width="0px"></asp:TextBox>
            <div class="clearage">
            </div>
            <div class="title-wave">
                Results:
            </div>
            <asp:TextBox ID="rtxtOutput" name="rtxtOutput" runat="server" AutoPostBack="true"
                Height="150px" OnTextChanged="rtxtOutput_TextChanged" TextMode="MultiLine" OnLoad="rtxtOutput_Load"></asp:TextBox>
            <div class="clearage">
            </div>
            <asp:ImageButton ID="ImageButton1" runat="server" OnClick="btnExecute_Click" ImageUrl="~/images/start.png"
                ToolTip="Run Script or Ctrl+Enter" />
            <asp:ImageButton ID="ImageButton2" runat="server" OnClick="btnClear_Click" ImageUrl="~/images/clear.png"
                ToolTip="Clear Script" />
            <asp:ImageButton ID="ImageButton3" runat="server" OnClientClick="handleEvent(event)" ImageUrl="~/images/variable.png"
                ToolTip="Click To Declare Variables" />
            <asp:ImageButton ID="ImageButton5" runat="server" OnClick="buttonSave_Click" ImageUrl="~/images/save.png"
                ToolTip="Save Script Locally" />
            <input type="hidden" id="hidValue" runat="server" />
            <input type="hidden" id="txtValue" runat="server" />
        </form>
        <asp:PlaceHolder ID="PlaceHolder1" runat="server">
            <script type="text/javascript">
                var editorDiv = document.getElementById('editor');
                editorDiv.ondragover = function () {
                    this.className = 'hover ace_editor ace-chrome';
                    return false;
                };
                editorDiv.ondragend = function () {
                    this.className = ' ace_editor ace-chrome';
                    return false;
                };
                editorDiv.ondrop = function (e) {
                    this.className = ' ace_editor ace-chrome';
                    e.preventDefault();
                
                    // vars
                    var file = e.dataTransfer.files[0],
                        reader = new FileReader();
                
                    // Debug: Display file name & attributes
                    // console.log('Name: ' + e.dataTransfer.files[0].name);
                    switch (true) {
                        case e.dataTransfer.files[0].name.indexOf('.txt') != -1:
                        case e.dataTransfer.files[0].name.indexOf('.vb') != -1:
                        case e.dataTransfer.files[0].name.indexOf('.vbs') != -1:
                            document.getElementById('<%= hidValue.ClientID %>').value = e.dataTransfer.files[0].name;
                            reader.onload = function (event) {
                                console.log(event.target);
                                editor.getSession();
                                editor.setValue(event.target.result, -1);
                                editor.focus();
                                clearTextBox();
                            };
                            console.log(file);
                            reader.readAsText(file);
                            return false;
                        default:
                            console.log('No Bueno');
                            alert('File Extension Not Supported!');
                            return true;
                    }
                };
            
                var editor = ace.edit("editor");            
                var erroneousLine;
                editor.$blockScrolling = Infinity;
                editor.session.setNewLineMode("windows");
                editor.setTheme("ace/theme/chrome");
                editor.getSession().setMode("ace/mode/vbscript");
                editor.setHighlightActiveLine(true);
                editor.setOption("showPrintMargin", false);
                editor.setOption("showInvisibles", true);
                var textarea = $('textarea[name="fctb"]');
                editor.getSession().setValue(textarea.val());
                editor.getSession().on('change', function () {
                    textarea.val(editor.getSession().getValue());
                });
                editor.setOptions({
                    fontFamily: "consolas",
                    fontSize  : "12px"
                });
                var keywords = (<%=passJS %>);                
            editor.findAll(keywords, {
                backwards    : true,
                wrap         : true,
                caseSensitive: true,
                range        : null,
                wholeWord    : true,
                regExp       : true
            });
            
            function clearTextBox() {
                document.getElementById('<%=rtxtOutput.ClientID%>').value = "";
            }

            $('body').keypress(function(event) {
                var keycode = (event.keyCode ? event.keyCode : event.which);
                if ((event.keyCode == 10 || event.keyCode == 13) && event.ctrlKey) {
                    $('#ImageButton1').click();
                }
            });

            function handleEvent(oEvent) {                
                var listString = document.getElementById('txtValue').value;
                var arrVars = listString.split('~');                
                for (i = 0; i < arrVars.length; i ++){
                    var strVar = "Dim " +  arrVars[i];
                    var session = editor.session;
                    if (session.doc.getValue().indexOf(strVar) == -1){
                        session.insert({
                            row: 0,
                            column: 0
                        },"Dim " +  arrVars[i] + " = \"" + "\"" + "\n");                        
                    }
                }
            }
            </script>
        </asp:PlaceHolder>
    </div>
</body>
</html>