Public Class JsonViewer

    ' Constructor to set additional form properties if needed
    'Public Sub New()
    ' This call is required by the designer.
    '   InitializeComponent()

    ' Set form properties
    'Me.Text = "JSON File Viewer"
    'Me.MaximizeBox = False
    'Me.MinimizeBox = True
    'Me.FormBorderStyle = FormBorderStyle.FixedDialog
    'Me.StartPosition = FormStartPosition.CenterParent
    'End Sub

    Public Sub New()
        Try
            InitializeComponent() ' Ensure this line runs successfully
            Me.Text = "JSON File Viewer"
            Me.MaximizeBox = False
            Me.MinimizeBox = True
            Me.FormBorderStyle = FormBorderStyle.FixedDialog
            Me.StartPosition = FormStartPosition.CenterParent
        Catch ex As Exception
            MessageBox.Show($"Error in constructor: {ex.Message}", "Constructor Error")
            Throw
        End Try
    End Sub


    Public Sub LoadJson(jsonText As String)
        Try
            ' Parse the JSON string
            Dim jsonObj = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
            ' Populate the TreeView with structured nodes
            TreeViewJson.Nodes.Clear()

            ' Create the root node
            Dim rootNode As New TreeNode("Root")
            TreeViewJson.Nodes.Add(rootNode)

            ' Add the Header node
            If jsonObj("Header") IsNot Nothing Then
                Dim headerNode As New TreeNode("Header")
                headerNode.Nodes.Add(New TreeNode(jsonObj("Header").ToString()))
                rootNode.Nodes.Add(headerNode)
            End If

            ' Create the Sections node
            If jsonObj("Sections") IsNot Nothing Then
                Dim sectionsNode As New TreeNode("Sections")
                rootNode.Nodes.Add(sectionsNode)

                ' Iterate over each section and add as a child of Sections
                Dim sections As Newtonsoft.Json.Linq.JObject = jsonObj("Sections")
                For Each section As KeyValuePair(Of String, Newtonsoft.Json.Linq.JToken) In sections
                    Dim sectionNode As New TreeNode(section.Key)
                    sectionsNode.Nodes.Add(sectionNode)
                    AddNode(section.Value, sectionNode)
                Next
            End If

            TreeViewJson.ExpandAll()
        Catch ex As Exception
            MessageBox.Show("Error parsing JSON: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub




    Private Sub AddNode(jToken As Newtonsoft.Json.Linq.JToken, treeNode As TreeNode)
        If jToken Is Nothing Then Return

        Select Case jToken.Type
            Case Newtonsoft.Json.Linq.JTokenType.Object
                Dim jsonObject As Newtonsoft.Json.Linq.JObject = CType(jToken, Newtonsoft.Json.Linq.JObject)
                For Each propertyPair As KeyValuePair(Of String, Newtonsoft.Json.Linq.JToken) In jsonObject
                    Dim childNode As New TreeNode(propertyPair.Key)
                    treeNode.Nodes.Add(childNode)
                    AddNode(propertyPair.Value, childNode)
                Next

            Case Newtonsoft.Json.Linq.JTokenType.Array
                Dim jsonArray As Newtonsoft.Json.Linq.JArray = CType(jToken, Newtonsoft.Json.Linq.JArray)
                For Each item As Newtonsoft.Json.Linq.JToken In jsonArray
                    If item.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                        ' Handle each object in the array, looking for "Index" and "Value"
                        Dim index As String = item("Index").ToString()
                        Dim value As String = item("Value").ToString()

                        ' Create a parent node for the index
                        Dim indexNode As New TreeNode("Index " & index)
                        treeNode.Nodes.Add(indexNode)

                        ' Add the value as a child node
                        Dim valueNode As New TreeNode("Value: " & value)
                        indexNode.Nodes.Add(valueNode)
                    End If
                Next

            Case Else
                ' For simple values, set the text directly
                treeNode.Text &= ": " & jToken.ToString()
        End Select
    End Sub





End Class
