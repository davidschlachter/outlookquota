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
    Public Shared theMessage As String
    'The default data path (AppData\EmailSizer)
    Public Shared RootPath As String = Environ("AppData") & "\EmailSizer"


    ' This function searches through subfolders
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

            Dim oForm As FirstRunProgress
            oForm = New FirstRunProgress

            Try
                ' Set the lock variable (now unnecessary because we only run at start)
                areWeRunning = True

                ' Set up the objects and variables
                Dim inb As Outlook.MAPIFolder, m As Outlook.MailItem, f As Outlook.MAPIFolder, olApp As Outlook.Application = New Outlook.Application
                inb = olApp.GetNamespace("mapi").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent

                ' Create the cache folder if it doesn't exist
                If Dir(RootPath, vbDirectory) = "" Then
                    Try
                        theMessage = "We'll be calculating the size of your mailbox for the first time. This may take a few minutes."
                        MkDir(RootPath)
                    Catch ex As System.Exception
                        Exit Sub
                    End Try
                Else
                    theMessage = "Currently updating your quota usage. This may take a few minutes."
                End If
                ' Make sure that the counters are at zero to start
                a = 0
                s = 0
                Dim b As Long = 0

                ' Update the progress bar with the current folder
                allFolders = inb.Folders.Count
                oForm.Show()

                ' Tally the sizes for the root of the mailbox (usually empty)
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

                ' Process all of the subfolders, and keep the progress bar up-to-date
                For Each f In inb.Folders
                    currentFolder = f.Name
                    atFolderCount = atFolderCount + 1
                    oForm.Refresh()
                    itmsiz(f)
                Next

                ' The results... (for display in the ribbon)
                NumberUsage = (a + s) / 1000000.0#
                PercentageQuota = ((a + s) / Quota) * 100
                RawSize = (a + s)

                ' Remove the lock
                areWeRunning = False

            Catch e As System.Exception
                ' In case of errors, remove the lock
                areWeRunning = False
                ' Then log what the error was
                Using errorwriter As StreamWriter = File.AppendText(RootPath & "\errorlog.txt")
                    errorwriter.WriteLine(e.ToString)
                    errorwriter.Close()
                End Using
            End Try

            ' Close the progress bar
            oForm.Close()
            oForm = Nothing

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

    'Enable the ribbon XML File
    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon1()
    End Function

End Class
