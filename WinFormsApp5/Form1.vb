Imports System.ComponentModel
Imports System.Formats.Asn1
Imports System.Net.NetworkInformation
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Threading

Public Class Form1
    ' Form-level variables
    Dim flagSort As Boolean
    Dim x, y As Integer
    Dim panel As New Panel
    Dim listSort As New ListBox
    Dim hashPattren As New Dictionary(Of Integer, String())
    Dim prefixe() As String
    Dim suffixe() As String
    Dim ignore As New Dictionary(Of String, Boolean)
    Dim newItems As New Dictionary(Of String, HashSet(Of String))

    ' Form Load Event
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize form size and position variables
        Me.Size = New Size(1025, 666)
        x = 12
        y = 12

        ' Load initial data
        ignore.insertIgnoreWords()
        hashPattren.insrtPattrens()
        prefixe.insrtPrefixe()
        suffixe.insrtSuffixe()
    End Sub

    ' Button Click Event: Start processing input
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim words As List(Of String) = TextBox1.Text.Split(" ").ToList
        root(words)
        newForm()
    End Sub

    ' Process input words to extract root
    Private Sub root(words As List(Of String))
        words.RemoveAll(Function(s) s = "")

        Dim flip As Boolean = True
        Dim ans As String = ""
        Dim edit As String

        For Each word As String In words
            word.ReplaceNotGoodChar()
            word = word.Trim()
            edit = word

            If word.isIgnore(ignore) Then
                ignore(word) = True
                Continue For
            End If

            If word.isDigit Then Continue For

            If word = "الله" Then
                insertInMap(word, word, word)
                Continue For
            End If

            firstRemove(edit, ans, flip)
            flip = True
            secondRemove(edit, ans, flip)
            insertInMap(edit, ans, word)

            ans = ""
            flip = True
        Next
    End Sub

    ' Remove prefixes from word
    Private Function deletePrefixe(ByRef edit As String) As Boolean
        For Each prefix As String In prefixe
            If edit.StartsWith(prefix) Then
                edit = edit.Substring(prefix.Length)
                Return True
            End If
        Next
        Return False
    End Function

    ' Remove suffixes from word
    Private Function deleteSuffixe(ByRef edit As String) As Boolean
        For Each suffix As String In suffixe
            If edit.EndsWith(suffix) Then
                edit = edit.Substring(0, edit.Length - suffix.Length)
                Return True
            End If
        Next
        Return False
    End Function

    ' Match word with patterns
    Private Function check(patterns As String(), word As String) As String
        For Each pattern As String In patterns
            Dim one = pattern.IndexOf("ف")
            Dim two = pattern.IndexOf("ع")
            Dim three = pattern.IndexOf("ل")

            Dim arr() As Char = word.ToCharArray()
            arr(one) = "ف"
            arr(two) = "ع"
            arr(three) = "ل"
            Dim transformedWord = New String(arr)

            If transformedWord = pattern Then
                Return word(one) + word(two) + word(three)
            End If
        Next
        Return ""
    End Function

    ' Remove prefixes
    Private Sub firstRemove(ByRef edit As String, ByRef ans As String, ByRef flip As Boolean)
        While ans = "" AndAlso flip
            If hashPattren.ContainsKey(edit.Length) Then
                ans = check(hashPattren(edit.Length), edit)
            End If

            If ans = "" Then
                flip = deletePrefixe(edit)
                If flip Then edit.editStartWithPrefixe()
            End If
        End While
    End Sub

    ' Remove suffixes
    Private Sub secondRemove(ByRef edit As String, ByRef ans As String, ByRef flip As Boolean)
        While ans = "" AndAlso flip
            If hashPattren.ContainsKey(edit.Length) Then
                ans = check(hashPattren(edit.Length), edit)
            End If

            If ans = "" Then
                flip = deleteSuffixe(edit)
                If flip Then edit.editEndWithSuffixe()
            End If
        End While
    End Sub

    ' Insert word into map
    Private Sub insertInMap(ByRef edit As String, ByRef ans As String, ByRef word As String)
        If edit.Length = 3 AndAlso ans = "" Then ans = edit
        If word.Length = 3 AndAlso ans = "" Then ans = word

        If ans.Length <> 3 Then Exit Sub

        If Not newItems.ContainsKey(ans) Then
            newItems.Add(ans, New HashSet(Of String))
        End If

        If Not newItems(ans).Contains(word) Then
            newItems(ans).Add(word)
            createButton(ans, word)
        End If
    End Sub

    ' Create new form for output
    Private Sub newForm()
        Me.Hide()
        Me.Size = New Size(1368, 769)
        Me.BackgroundImage = New Bitmap(Button1.BackgroundImage)
        Thread.Sleep(300)

        Me.Controls.Remove(Button1)
        Me.Controls.Remove(TextBox1)
        Me.Controls.Remove(Label1)

        panel.AutoSize = True
        CheckBox1.Visible = True

        btnSort()
        createBox(panel)

        Me.Show()
        buttonForText()
    End Sub

    ' Create sort button
    Private Sub btnSort()
        Dim btn As New Button With {
            .Size = New Size(131, 40),
            .Text = "Sort",
            .Tag = "Sort",
            .Name = "Sort",
            .Font = New Font(Button1.Font.FontFamily, Button1.Font.Size, Button1.Font.Style),
            .Location = New Point(990, 500),
            .BackgroundImage = New Bitmap(Button1.BackgroundImage)
        }

        AddHandler btn.Click, AddressOf sort_btn
        Me.Controls.Add(btn)

        createBox(listSort)
        listSort.RightToLeft = RightToLeft.Yes
        listSort.Visible = False
    End Sub

    ' Create buttons dynamically
    Private Sub createButton(ans As String, word As String)
        Dim btn As New Button With {
            .Size = New Size(131, 40),
            .Location = New Point(x, y),
            .Text = word,
            .Tag = ans
        }

        AddHandler btn.MouseHover, AddressOf hover_btn
        AddHandler btn.MouseLeave, AddressOf hover_btn

        x += 129
        If x >= 790 Then
            y += 46
            x = 12
        End If

        panel.Controls.Add(btn)
    End Sub

    ' Create container for buttons or lists
    Private Sub createBox(ByRef container As Object)
        container.Size = New Size(792, 681)
        container.Location = New Point(2, 12)
        container.BorderStyle = BorderStyle.Fixed3D
        container.BackgroundImage = New Bitmap(Button1.BackgroundImage)
        Me.Controls.Add(container)
    End Sub

    ' Button hover behavior
    Private Sub hover_btn(sender As Object, e As EventArgs) Handles Button1.MouseLeave
        Dim btn As Button = CType(sender, Button)
        Dim newName = btn.Text
        btn.Text = btn.Tag
        btn.Tag = newName
    End Sub

    ' Add new text button
    Private Sub buttonForText()
        Dim btn As New Button With {
            .Size = New Size(131, 40),
            .Location = New Point(980, 280),
            .BackgroundImage = New Bitmap(My.Resources._6921087_ai),
            .Text = "Add New Text",
            .Tag = "Add New Text",
            .Name = "btnText",
            .AutoSize = True
        }

        AddHandler btn.MouseHover, AddressOf hover_addText
        AddHandler btn.Click, AddressOf click_addText
        Me.Controls.Add(btn)
    End Sub

    ' Hover behavior for Add New Text button
    Private Sub hover_addText(sender As Object, e As EventArgs) Handles Button1.MouseHover
        Dim btn As Button = CType(sender, Button)
        If btn.Name = "btnText" Then TextBox2.Visible = True
    End Sub

    ' Click event for Add New Text button
    Private Sub click_addText(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        If btn.Name = "btnText" Then
            TextBox2.Visible = False
            root(TextBox2.Text.Split(" ").ToList)
            TextBox2.Text = Nothing
        End If
    End Sub

    ' Sort button behavior
    Private Sub sort_btn(sender As Object, e As EventArgs) Handles Button1.Click
        Dim btn As Button = CType(sender, Button)
        If btn.Name = "Sort" Then
            btn.Text = If(flagSort, "Sort", "Un Sort")
            panel.Visible = Not panel.Visible
            listSort.Visible = Not listSort.Visible
            flagSort = Not flagSort
            printRoot()
        End If
    End Sub

    ' Display root words
    Private Sub printRoot()
        If Not panel.Visible Then
            listSort.Items.Clear()
            For Each item In newItems
                Dim setWords = String.Join(" , ", item.Value)
                Dim arr() = setWords.Split(" , ")

                If arr.Length <= 8 Then
                    listSort.Items.Add($"{item.Key} = [ {setWords} ]")
                Else
                    Dim firstHalf = String.Join(" , ", arr.Take(8))
                    Dim secondHalf = String.Join(" , ", arr.Skip(8))
                    listSort.Items.Add($"{item.Key} = [ {firstHalf}")
                    listSort.Items.Add($"{secondHalf} ]")
                End If
            Next
        End If
    End Sub

    ' Check box for ignored words
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            Thread.Sleep(200)
            listSort.Items.Clear()
            For Each item In ignore
                If item.Value Then listSort.Items.Add(item.Key)
            Next
        Else
            printRoot()
        End If
    End Sub
End Class
