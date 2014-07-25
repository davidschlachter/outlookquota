Imports System.IO
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Tools.Outlook
Imports Microsoft.Office.Core

Public Class ThisAddIn
    ' Set the file-size counters globally so that they don't reset!
    Public s As Long
    Public a As Long
    'Permit ribbon manipulation
    Public ribbon As Office.IRibbonUI
    'The quota in bytes (and other quota variables that the ribbon will need for display)
    Public Shared Quota As Double = 2000000000
    Public Shared NumberUsage As Integer
    Public Shared PercentageQuota As Integer
    'For evaluating the contents of 'Deleted Items'
    Public Shared FirstIDDelItems As String
    Public Shared NumDelItems As Integer
    'The 'are-we-running' variable
    Public Shared areWeRunning As Boolean
    'The default data path (AppData\EmailSizer)
    Public Shared RootPath As String = Environ("AppData") & "\EmailSizer"

    ' This function loops through all folders and calculates the total size of the root-level mailbox
    Public Sub folsize()
        'If this fails, stop, and let go of the lock  :)
        Try
            'Lock-file style variable: if it's one, don't execute; when we finish, set to 0; if we fail, set to 0.
            areWeRunning = 1

            ' CAUTION: What about non-MailItem items???
            Dim inb As Outlook.MAPIFolder, m As Outlook.MailItem, f As Outlook.MAPIFolder, olApp As Outlook.Application = New Outlook.Application
            inb = olApp.GetNamespace("mapi").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent

            ' Check if the cache directory exists
            If Dir(RootPath, vbDirectory) = "" Then
                MkDir(RootPath)
                MsgBox("We'll be calculating the size of your inbox for the first time. This may take some time. Please be patient  :)")
            End If

            ' Zero the size variables (just in case  :)
            a = 0
            s = 0
            Dim b As Long = 0

            ' Check for any items at the root (unlikely)
            For Each m In inb.Items
                a = a + m.Size
            Next

            ' Now, loop through all the folders
            For Each f In inb.Folders
                itmsiz(f)
            Next

            'Compute the results
            NumberUsage = (a + s) / 1000000.0#
            PercentageQuota = ((a + s) / Quota) * 100
            a = 0
            s = 0

            'Let go of the lock
            areWeRunning = 0

        Finally
            'On error...
            areWeRunning = 0
        End Try
    End Sub




    ' This function searches through subfolders (part of folsize)
    Public Sub itmsiz(ByVal fol As Object)

        Dim f As Object, count As Integer, firstItem As String, size As Long, b As Long

        ' Check for any more subfolders
        For Each f In fol.Folders
            itmsiz(f)
        Next

        ' Actually calculate the sizes
        ' First, check if we have an ID-based folder, if not, make it
        Dim ItemFolder As String = RootPath & "\" & fol.EntryID
        If Dir(ItemFolder, vbDirectory) = "" Then
            MkDir(ItemFolder)
        End If

        ' Quick error prevention check -- if the current count is zero, get rid of the cache
        If fol.Items.count = 0 Then
            If Not Dir(ItemFolder & "\Count", vbDirectory) = "" Then
                SetAttr(ItemFolder & "\Count", vbNormal)
                Kill(ItemFolder & "\Count")
            End If
        End If

        ' Do we have an item count in the folder? Does it match?
        If Dir(ItemFolder & "\Count", vbDirectory) = "" Then
            'No count found -- creating it
            Dim countwriter As New System.IO.StreamWriter(ItemFolder & "\Count")
            countwriter.Write(fol.Items.count)
            countwriter.Close()
            ' Also creating the first item ID
            If fol.Items.count > 0 Then
                Dim idwriter As New System.IO.StreamWriter(ItemFolder & "\firstItem")
                idwriter.Write(fol.Items.Item(1).EntryID)
                idwriter.Close()
            End If
            ' Then, crunch the numbers -- there was no cache
            b = 0
            For Each m In fol.Items
                b = b + m.size
            Next
            s = s + b
            Dim sizewriter As New System.IO.StreamWriter(ItemFolder & "\Size")
            sizewriter.Write(b)
            sizewriter.Close()
            b = 0
        Else ' If we did have a cache with the count, get the count
            Dim countreader As New System.IO.StreamReader(ItemFolder & "\Count")
            count = countreader.ReadLine()
            countreader.Close()
            ' Now, compare the count
            If count = fol.Items.count Then
                ' They're equal: further testing required
                If Dir(ItemFolder & "\firstItem", vbDirectory) = "" Then
                    ' If there was no firstItem ID, create it, then crunch numbers
                    If fol.Items.count > 0 Then
                        Dim firstwriter As New System.IO.StreamWriter(ItemFolder & "\firstItem")
                        firstwriter.Write(fol.Items.Item(1).EntryID)
                        firstwriter.Close()
                    End If
                    ' Then, crunch the numbers -- there was no cache
                    b = 0
                    For Each m In fol.Items
                        b = b + m.size
                    Next
                    s = s + b
                    Dim sizewriter As New System.IO.StreamWriter(ItemFolder & "\Size")
                    sizewriter.Write(b)
                    sizewriter.Close()
                    b = 0
                Else ' We did have a firstItem ID, let's check if they're the same...
                    Dim firstreader As New System.IO.StreamReader(ItemFolder & "\firstItem")
                    firstItem = firstreader.ReadLine()
                    firstreader.Close()
                    ' Now, compare the firstItem IDs
                    If firstItem = fol.Items.Item(1).EntryID Then
                        'They're the same -- i.e. the folder hasn't changed
                        'Need to check now if we have a cached value
                        If Dir(ItemFolder & "\Size", vbDirectory) = "" Then
                            'No file; crunch numbers, make size file
                            b = 0
                            For Each m In fol.Items
                                b = b + m.size
                            Next
                            s = s + b
                            Dim sizewriter As New System.IO.StreamWriter(ItemFolder & "\Size")
                            sizewriter.Write(b)
                            sizewriter.Close()
                            b = 0
                        Else
                            'We have the file -- read numbers
                            Dim sizereader As New System.IO.StreamReader(ItemFolder & "\Size")
                            size = sizereader.ReadLine()
                            sizereader.Close()
                            s = s + size
                        End If
                    Else
                        'The folder has changed. Time to get a new first item ID, then calculate!
                        If fol.Items.count > 0 Then
                            Dim firstwriter As New System.IO.StreamWriter(ItemFolder & "\firstItem")
                            firstwriter.Write(fol.Items.Item(1).EntryID)
                            firstwriter.Close()
                        End If
                        ' Then, crunch the numbers -- there was no cache
                        b = 0
                        For Each m In fol.Items
                            b = b + m.size
                        Next
                        s = s + b
                        Dim sizewriter As New System.IO.StreamWriter(ItemFolder & "\Size")
                        sizewriter.Write(b)
                        sizewriter.Close()
                        b = 0
                    End If
                End If
            Else
                ' They're different, so we'll just write the new count (and first item ID) and tally up the sizes
                Dim countwriter As New System.IO.StreamWriter(ItemFolder & "\Count")
                countwriter.Write(fol.Items.count)
                countwriter.Close()
                If fol.Items.count > 0 Then
                    Dim firstwriter As New System.IO.StreamWriter(ItemFolder & "\firstItem")
                    firstwriter.Write(fol.Items.Item(1).EntryID)
                    firstwriter.Close()
                End If
                b = 0
                For Each m In fol.Items
                    b = b + m.size
                Next
                s = s + b
                Dim sizewriter As New System.IO.StreamWriter(ItemFolder & "\Size")
                sizewriter.Write(b)
                sizewriter.Close()
                b = 0
            End If
        End If
    End Sub







    'Run on start  :)
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        'At start, make sure we don't have the lock
        areWeRunning = 0
        'Run the sizer
        folsize()
    End Sub

    'Update the counts when new mail arrives
    Private Sub OutLook_NewMaiItem() Handles Application.NewMail
        If areWeRunning = 0 Then
            folsize()
            ribbon.Invalidate()
        End If
    End Sub

    'And, detect when users empty Deleted Items, and also update
    Private Sub OutLook_ItemLoad() Handles Application.ItemLoad
        Dim delitms As Outlook.MAPIFolder, olApp As Outlook.Application = New Outlook.Application
        delitms = olApp.GetNamespace("mapi").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems)
        If areWeRunning = 0 Then
            If NumDelItems = delitms.Items.Count Then
                If delitms.Items.Count <> 0 Then
                    If FirstIDDelItems = delitms.Items.Item(1).EntryID Then
                        Exit Sub
                    Else
                        folsize()
                        FirstIDDelItems = delitms.Items.Item(1).EntryID
                        ribbon.Invalidate()
                    End If
                Else
                    Exit Sub
                End If
            Else
                folsize()
                NumDelItems = delitms.Items.Count
                If delitms.Items.Count <> 0 Then
                    FirstIDDelItems = delitms.Items.Item(1).EntryID
                End If
                ribbon.Invalidate()
            End If
        End If
    End Sub

    'Enable the ribbon XML File
    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon1()
    End Function

End Class
