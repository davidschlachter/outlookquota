Imports System.IO
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Tools.Outlook
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices

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
    Public Shared RawSize As Long
    'For evaluating the contents of 'Deleted Items'
    Public Shared FirstIDDelItems As String
    Public Shared NumDelItems As Integer
    'The 'are-we-running' variable
    Public Shared areWeRunning As Boolean
    'For the progress bar
    Public Shared allFolders As Integer
    Public Shared atFolderCount As Integer
    Public Shared currentFolder As String
    'The default data path (AppData\EmailSizer)
    Public Shared RootPath As String = Environ("AppData") & "\EmailSizer"

    ' This function loops through all folders and calculates the total size of the root-level mailbox
    Public Sub folsize()
        'If this fails, stop, and let go of the lock  :)
        Try
            'Lock-file style variable: if it's one, don't execute; when we finish, set to 0; if we fail, set to 0.
            areWeRunning = True

            ' CAUTION: What about non-MailItem items???
            Dim inb As Outlook.MAPIFolder, m As Outlook.MailItem, f As Outlook.MAPIFolder, olApp As Outlook.Application = New Outlook.Application
            inb = olApp.GetNamespace("mapi").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent

            ' Check if the cache directory exists
            If Dir(RootPath, vbDirectory) = "" Then
                MkDir(RootPath)
                MsgBox("We'll be calculating the size of your inbox for the first time. This may take a few minutes.")
            End If

            ' Zero the size variables (just in case  :)
            a = 0
            s = 0
            Dim b As Long = 0

            ' Get the total number of root-level folders
            allFolders = inb.Folders.Count

            ' Check for any items at the root (unlikely)

            For Each m In inb.Items
                Try
                    a = a + m.Size
                Catch e As System.Exception
                    Try
                        Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
                            errorwriter.WriteLine(e.ToString)
                            errorwriter.Close()
                        End Using
                    Catch exc As System.Exception
                        MsgBox("Unable to write error log: " & exc.ToString)
                    End Try
                Finally
                    Marshal.ReleaseComObject(m)
                End Try
            Next

            ' Now, loop through all the folders
            For Each f In inb.Folders
                itmsiz(f)
            Next

            'Compute the results
            NumberUsage = (a + s) / 1000000.0#
            PercentageQuota = ((a + s) / Quota) * 100
            RawSize = (a + s)
            a = 0
            s = 0

            'Let go of the lock
            areWeRunning = False

        Catch e As System.Exception
            'On error...
            areWeRunning = False
            Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
                errorwriter.WriteLine(e.ToString)
                errorwriter.Close()
            End Using
        End Try
    End Sub




    ' This function searches through subfolders (part of folsize)
    Public Sub itmsiz(ByVal fol As Object)
        Try
            Dim f As Object, count As Integer, firstItem As String, size As Long, b As Long

            'Make sure that the variable is set!
            areWeRunning = True

            ' Check for any more subfolders
            For Each f In fol.Folders
                itmsiz(f)
            Next

            ' Actually calculate the sizes
            ' First, check if we have an ID-based folder, if not, make it
            Dim ItemFolder As String = RootPath & "\" & fol.EntryID
            Try
                If Dir(ItemFolder, vbDirectory) = "" Then
                    MkDir(ItemFolder)
                End If
            Catch e As System.Exception
                Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
                    errorwriter.WriteLine(e.ToString)
                    errorwriter.Close()
                End Using
            End Try


            ' Quick error prevention check -- if the current count is zero, get rid of the cache
            Try
                If fol.Items.count = 0 Then
                    If Not Dir(ItemFolder & "\Count", vbDirectory) = "" Then
                        SetAttr(ItemFolder & "\Count", vbNormal)
                        Kill(ItemFolder & "\Count")
                    End If
                End If
            Catch e As System.Exception
                Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
                    errorwriter.WriteLine(e.ToString)
                    errorwriter.Close()
                End Using
            End Try

            ' Do we have an item count in the folder? Does it match?
            If Dir(ItemFolder & "\Count", vbDirectory) = "" Then
                'No count found -- creating it
                Dim countwriter As New StreamWriter(ItemFolder & "\Count")
                countwriter.Write(fol.Items.count)
                countwriter.Close()
                ' Also creating the first item ID
                If fol.Items.count > 0 Then
                    Dim idwriter As New StreamWriter(ItemFolder & "\firstItem")
                    idwriter.Write(fol.Items.Item(1).EntryID)
                    idwriter.Close()
                End If
                ' Then, crunch the numbers -- there was no cache
                b = 0

                For Each m In fol.Items
                    Try
                        b = b + m.size
                    Catch e As System.Exception
                        Try
                            Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
                                errorwriter.WriteLine(e.ToString)
                                errorwriter.Close()
                            End Using
                        Catch exc As System.Exception
                            MsgBox("Unable to write error log: " & exc.ToString)
                        End Try
                    Finally
                        Marshal.ReleaseComObject(m)
                    End Try
                Next

                s = s + b

                Dim sizewriter As New StreamWriter(ItemFolder & "\Size")
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
                            Dim firstwriter As New StreamWriter(ItemFolder & "\firstItem")
                            firstwriter.Write(fol.Items.Item(1).EntryID)
                            firstwriter.Close()
                        End If
                        ' Then, crunch the numbers -- there was no cache
                        b = 0
                        For Each m In fol.Items
                            Try
                                b = b + m.size
                            Catch e As System.Exception
                                Try
                                    Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
                                        errorwriter.WriteLine(e.ToString)
                                        errorwriter.Close()
                                    End Using
                                Catch exc As System.Exception
                                    MsgBox("Unable to write error log: " & exc.ToString)
                                End Try
                            Finally
                                Marshal.ReleaseComObject(m)
                            End Try
                        Next
                        s = s + b
                        Dim sizewriter As New StreamWriter(ItemFolder & "\Size")
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
                                    Try
                                        b = b + m.size
                                    Catch e As System.Exception
                                        Try
                                            Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
                                                errorwriter.WriteLine(e.ToString)
                                                errorwriter.Close()
                                            End Using
                                        Catch exc As System.Exception
                                            MsgBox("Unable to write error log: " & exc.ToString)
                                        End Try
                                    Finally
                                        Marshal.ReleaseComObject(m)
                                    End Try
                                Next
                                s = s + b
                                Dim sizewriter As New StreamWriter(ItemFolder & "\Size")
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
                                Dim firstwriter As New StreamWriter(ItemFolder & "\firstItem")
                                firstwriter.Write(fol.Items.Item(1).EntryID)
                                firstwriter.Close()
                            End If
                            ' Then, crunch the numbers -- there was no cache
                            b = 0
                            For Each m In fol.Items
                                Try
                                    b = b + m.size
                                Catch e As System.Exception
                                    Try
                                        Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
                                            errorwriter.WriteLine(e.ToString)
                                            errorwriter.Close()
                                        End Using
                                    Catch exc As System.Exception
                                        MsgBox("Unable to write error log: " & exc.ToString)
                                    End Try
                                Finally
                                    Marshal.ReleaseComObject(m)
                                End Try
                            Next
                            s = s + b
                            Dim sizewriter As New StreamWriter(ItemFolder & "\Size")
                            sizewriter.Write(b)
                            sizewriter.Close()
                            b = 0
                        End If
                    End If
                Else
                    ' They're different, so we'll just write the new count (and first item ID) and tally up the sizes
                    Dim countwriter As New StreamWriter(ItemFolder & "\Count")
                    countwriter.Write(fol.Items.count)
                    countwriter.Close()
                    If fol.Items.count > 0 Then
                        Dim firstwriter As New StreamWriter(ItemFolder & "\firstItem")
                        firstwriter.Write(fol.Items.Item(1).EntryID)
                        firstwriter.Close()
                    End If
                    b = 0
                    For Each m In fol.Items
                        Try
                            b = b + m.size
                        Catch e As System.Exception
                            Try
                                Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
                                    errorwriter.WriteLine(e.ToString)
                                    errorwriter.Close()
                                End Using
                            Catch exc As System.Exception
                                MsgBox("Unable to write error log: " & exc.ToString)
                            End Try
                        Finally
                            Marshal.ReleaseComObject(m)
                        End Try
                    Next
                    s = s + b
                    Dim sizewriter As New StreamWriter(ItemFolder & "\Size")
                    sizewriter.Write(b)
                    sizewriter.Close()
                    b = 0
                End If
            End If
        Catch e As System.Exception
            Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
                errorwriter.WriteLine(e.ToString)
                errorwriter.Close()
            End Using
        End Try

    End Sub







    'Run on start  :)
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Try
            'At start, make sure we don't have the lock
            areWeRunning = False

            'Try with progress bar for first run
            If Dir(RootPath, vbDirectory) = "" Then
                MkDir(RootPath)
                Dim oForm As FirstRunProgress
                oForm = New FirstRunProgress


                ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
                ' Here we're pasting the folsize code, and modifiying it to update the item count while it runs '
                ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' 
                Try
                    areWeRunning = True
                    Dim inb As Outlook.MAPIFolder, m As Outlook.MailItem, f As Outlook.MAPIFolder, olApp As Outlook.Application = New Outlook.Application
                    inb = olApp.GetNamespace("mapi").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent
                    If Dir(RootPath, vbDirectory) = "" Then
                        MkDir(RootPath)
                    End If
                    a = 0
                    s = 0
                    Dim b As Long = 0
                    allFolders = inb.Folders.Count
                    oForm.Show()
                    For Each m In inb.Items
                        Try
                            atFolderCount = 1
                            currentFolder = inb.Name
                            oForm.Refresh()
                            a = a + m.Size
                        Catch e As System.Exception
                            Try
                                Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
                                    errorwriter.WriteLine(e.ToString)
                                    errorwriter.Close()
                                End Using
                            Catch exc As System.Exception
                                MsgBox("Unable to write error log: " & exc.ToString)
                            End Try
                        Finally
                            Marshal.ReleaseComObject(m)
                        End Try
                    Next

                    For Each f In inb.Folders
                        currentFolder = f.Name
                        atFolderCount = atFolderCount + 1
                        oForm.Refresh()
                        itmsiz(f)
                    Next

                    NumberUsage = (a + s) / 1000000.0#
                    PercentageQuota = ((a + s) / Quota) * 100
                    RawSize = (a + s)
                    a = 0
                    s = 0
                    areWeRunning = False
                Catch e As System.Exception
                    areWeRunning = False
                    Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
                        errorwriter.WriteLine(e.ToString)
                        errorwriter.Close()
                    End Using
                End Try
                ' ' ' ' ' ' ' ' ' ' ' '
                ' End of folsize code '
                ' ' ' ' ' ' ' ' ' ' ' ' 

                oForm.Close()
                oForm = Nothing
            Else
                folsize()
            End If
            'Remove the lock
            areWeRunning = False
        Catch e As System.Exception
            areWeRunning = False
            Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
                errorwriter.WriteLine(e.ToString)
                errorwriter.Close()
            End Using
        End Try
    End Sub

    'Update the counts when new mail arrives
    'Private Sub OutLook_NewMaiItem() Handles Application.NewMail
    '    Try
    '        If Not areWeRunning Then
    '            folsize()
    '            ribbon.Invalidate()
    '            areWeRunning = False
    '        Else
    '            Exit Sub
    '        End If
    '    Catch e As System.Exception
    '        'DEBUG code
    '        MsgBox(e.ToString)
    '        MsgBox("The error was in the new mail event handler")
    '    End Try
    'End Sub

    'And, detect when users empty Deleted Items, and also update
    'Private Sub OutLook_ItemLoad() Handles Application.ItemLoad
    '    Try
    '        Dim delitms As Outlook.MAPIFolder, olApp As Outlook.Application = New Outlook.Application
    '        delitms = olApp.GetNamespace("mapi").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems)
    '        If Not areWeRunning Then
    '            If NumDelItems = delitms.Items.Count Then
    '                If delitms.Items.Count <> 0 Then
    '                    If FirstIDDelItems = delitms.Items.Item(1).EntryID Then
    '                        Exit Sub
    '                    Else
    '                        folsize()
    '                        FirstIDDelItems = delitms.Items.Item(1).EntryID
    '                        ribbon.Invalidate()
    '                        areWeRunning = False
    '                    End If
    '                Else
    '                    Exit Sub
    '                End If
    '            Else
    '                folsize()
    '                NumDelItems = delitms.Items.Count
    '                If delitms.Items.Count <> 0 Then
    '                    FirstIDDelItems = delitms.Items.Item(1).EntryID
    '                End If
    '                Try
    '                    ribbon.Invalidate()
    '                Catch e As System.Exception
    '                End Try
    '                areWeRunning = False
    '            End If
    '        Else
    '            Exit Sub
    '        End If
    '    Catch e As System.Exception
    '        'DEBUG code
    '        MsgBox(e.ToString)
    '        MsgBox("The error was in the load item handler")
    '    End Try
    'End Sub

    'Enable the ribbon XML File
    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon1()
    End Function

End Class
