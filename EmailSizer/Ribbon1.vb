Imports System.Windows.Forms

'TODO:  Follow these steps to enable the Ribbon (XML) item:

'2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
'   actions, such as clicking a button. Note: if you have exported this Ribbon from the
'   Ribbon designer, move your code from the event handlers to the callback methods and
'   modify the code to work with the Ribbon extensibility (RibbonX) programming model.

'3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.

'For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

<Runtime.InteropServices.ComVisible(True)> _
Public Class Ribbon1
    Implements Office.IRibbonExtensibility
    Private ribbon As Office.IRibbonUI
    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("EmailSizer.Ribbon1.xml")
    End Function

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1.
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
        'Via 'Invalidating Ribbon from Outside Ribbon' on MSDN Social...
        Globals.QuotaTool.ribbon = Me.ribbon
    End Sub


    'This dynamically updates the button text (and sets a condition for if we haven't sized yet)
    Public Function get_LabelName(ByVal control As Office.IRibbonControl) As String
        If QuotaTool.RawSize = 0 Then
            Return "Click me to update!"
        Else
            Return QuotaTool.PercentageQuota & " Percent Used"
        End If
    End Function

    'This dynamically updates the ribbon image based on the percentage of used quota
    Public Function showthebox(ByVal control As Office.IRibbonControl) As System.Drawing.Bitmap
        If QuotaTool.PercentageQuota < 75 Then
            Return My.Resources.Resource1.green
        ElseIf (QuotaTool.PercentageQuota >= 75 And QuotaTool.PercentageQuota < 90) Then
            Return My.Resources.Resource1.yellow
        ElseIf QuotaTool.PercentageQuota >= 90 Then
            Return My.Resources.Resource1.red
        Else
            'If all else fails, make the button yellow :)  (alert!)
            Return My.Resources.Resource1.yellow
        End If
    End Function

    'Action on click -- display detailed statistics
    Public Sub clickthebutton(ByVal control As Office.IRibbonControl)
        ribbon.InvalidateControl("QuotaIconButton")
        Dim detailsBox As New Details
        detailsBox.Show()
    End Sub

#End Region

#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
