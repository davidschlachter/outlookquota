Imports System.IO
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Tools.Outlook
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices

Public Class QuotaTool
    ' Set the file-size counters globally so that they don't reset!
    Public s As Long
    Public a As Long
    Public ItemFolder As String
    'Permit ribbon manipulation
    Public ribbon As Office.IRibbonUI
    'The quota in bytes (and other quota variables that the ribbon will need for display)
    Public Shared Quota As Double = 2000000000
    Public Shared NumberUsage As Integer
    Public Shared PercentageQuota As Integer
    Public Shared RawSize As Long
    'For evaluating the contents of 'Deleted Items'
    Public Shared FirstIDDelItems As String
    Public Shared NumDelItems As Integer
    'For the progress bar
    Public Shared allFolders As Integer
    Public Shared atFolder As Integer
    Public Shared currentFolder As String
    Public Shared theMessage As String
    'The default data path (AppData\EmailSizer)
    Public Shared RootPath As String = Environ("AppData") & "\EmailSizer"


    Private Sub QuotaTool_Startup() Handles Me.Startup
        mailboxsize()
    End Sub

    Public Sub mailboxsize()
        Try
            Dim m As Outlook.MailItem,
                f As Outlook.MAPIFolder,
                olApp As Outlook.Application = New Outlook.Application,
                inb As Outlook.MAPIFolder = olApp.GetNamespace("mapi").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent,
                progressForm As New Progress

            ' Create the cache folder if it doesn't exist
            If Dir(RootPath, vbDirectory) = "" Then
                Try
                    theMessage = "We'll be calculating the size of your mailbox for the first time. This may take a few minutes."
                    MkDir(RootPath)
                Catch ex As System.Exception
                    MsgBox("Unable to create the folder """ & RootPath & """.")
                    Exit Sub
                End Try
            Else
                theMessage = "Currently updating your quota usage. This may take a few minutes."
            End If

            ' Update the progress bar with the current folder
            allFolders = inb.Folders.Count
            progressForm.Show()

            ' Tally the sizes for the root of the mailbox (usually empty)
            For Each m In inb.Items
                Try
                    atFolder = 1
                    currentFolder = inb.Name
                    progressForm.Refresh()
                    a += m.Size
                Catch e As System.Exception
                    writeErrorLog(e)
                Finally
                    Marshal.ReleaseComObject(m)
                End Try
            Next

            ' Process all of the subfolders, and keep the progress bar up-to-date
            For Each f In inb.Folders
                currentFolder = f.Name
                atFolder = atFolder + 1
                progressForm.Refresh()
                itemsize(f)
            Next

            ' The results... (for display in the ribbon)
            NumberUsage = (a + s) / 1000000.0#
            PercentageQuota = ((a + s) / Quota) * 100
            RawSize = (a + s)

            progressForm.Close()

        Catch e As System.Exception
            writeErrorLog(e)
        End Try
    End Sub


    Public Sub itemsize(ByVal fol As Object)
        Dim f As Object,
            count As Integer,
            firstItem As String,
            size As Long

        For Each f In fol.Folders
            itemsize(f)
        Next

        Try
            ItemFolder = RootPath & "\" & fol.EntryID
            Try
                If Dir(ItemFolder, vbDirectory) = "" Then
                    MkDir(ItemFolder)
                End If
            Catch e As System.Exception
                writeErrorLog(e)
                Exit Sub
            End Try

            If fol.Items.count = 0 Then 'Delete cache if folder is empty
                If Not Dir(ItemFolder & "\Count", vbDirectory) = "" Then
                    SetAttr(ItemFolder & "\Count", vbNormal)
                    Kill(ItemFolder & "\Count")
                End If
            End If

            If Dir(ItemFolder & "\Count", vbDirectory) = "" Then 'We have no cache
                writeCountCache(fol)
                writeFirstItemID(fol)
                tally(fol)
            Else 'We have an item count
                Dim countreader As New System.IO.StreamReader(ItemFolder & "\Count")
                count = countreader.ReadLine()
                countreader.Close()
                If count = fol.Items.count Then 'Item count hasn't changed
                    If Dir(ItemFolder & "\firstItem", vbDirectory) = "" Then 'We didn't have a first item ID cache
                        writeFirstItemID(fol)
                        tally(fol)
                    Else ' We did have a firstItem ID
                        Dim firstreader As New System.IO.StreamReader(ItemFolder & "\firstItem")
                        firstItem = firstreader.ReadLine()
                        firstreader.Close()
                        If firstItem = fol.Items.Item(1).EntryID Then 'The folder hasn't changed
                            Try
                                Dim sizereader As New System.IO.StreamReader(ItemFolder & "\Size")
                                size = sizereader.ReadLine()
                                sizereader.Close()
                                s = s + size
                            Catch ex As System.Exception
                                tally(fol)
                            End Try
                        Else 'Same count, different items
                            writeFirstItemID(fol)
                            tally(fol)
                        End If
                    End If
                Else 'The item count changed
                    writeCountCache(fol)
                    writeFirstItemID(fol)
                    tally(fol)
                End If
            End If
        Catch e As System.Exception
            writeErrorLog(e)
        End Try
    End Sub

    Public Sub tally(ByVal fol As Object)
        Dim b As Long = 0
        For Each m In fol.Items
            Try
                b += m.size
            Catch e As System.Exception
                writeErrorLog(e)
            Finally
                Marshal.ReleaseComObject(m)
            End Try
        Next

        s += b

        Dim sizewriter As New StreamWriter(ItemFolder & "\Size")
        sizewriter.Write(b)
        sizewriter.Close()
    End Sub

    Public Sub writeErrorLog(ByVal e As System.Exception)
        Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
            errorwriter.WriteLine(e.ToString)
            errorwriter.Close()
        End Using
    End Sub

    Public Sub writeCountCache(ByVal fol As Object)
        Dim countwriter As New StreamWriter(ItemFolder & "\Count")
        countwriter.Write(fol.Items.count)
        countwriter.Close()
    End Sub

    Public Sub writeFirstItemID(ByVal fol As Object)
        If fol.Items.count > 0 Then
            Dim idwriter As New StreamWriter(ItemFolder & "\firstItem")
            idwriter.Write(fol.Items.Item(1).EntryID)
            idwriter.Close()
        End If
    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon1()
    End Function

End Class
