Imports System.ComponentModel
Imports System.Formats.Asn1
Imports System.Net.NetworkInformation
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Threading


Public Class Form1
    Dim flagSort As Boolean
    Dim x, y As Integer
    Dim panel As New Panel
    Dim listSort As New ListBox
    Dim hashPattren As New Dictionary(Of Integer, String())
    Dim prefixe() As String
    Dim suffixe() As String
    Dim ignore As New Dictionary(Of String, Boolean)
    Dim newItems As New Dictionary(Of String, HashSet(Of String))

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Size = New Size(1025, 666)
        x = 12
        y = 12
        ignore.insertIgnoreWords()
        hashPattren.insrtPattrens()
        prefixe.insrtPrefixe()
        suffixe.insrtSuffixe()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim words As New List(Of String)

        words = TextBox1.Text.Split(" ").ToList
        root(words)

        newForm()
    End Sub

    Private Sub root(words As List(Of String))

        words.RemoveAll(Function(s As String)
                            Return s = ""
                        End Function)

        Dim flip As Boolean = True
        Dim ans As String = ""
        Dim edit As String

        For Each i As String In words

            i.ReplaceNotGoodChar()
            'i.removeSpace()
            i = i.Trim()
            edit = i

            If i.isIgnore(ignore) Then
                ignore(i) = True
                Continue For
            End If

            If i.isDigit Then
                Continue For
            End If

            If i = "الله" Then
                insertInMap(i, i, i)
                Continue For
            End If
            firstRemove(edit, ans, flip)

            flip = True

            secondRemove(edit, ans, flip)


            insertInMap(edit, ans, i)

            ans = ""
            flip = True
        Next
    End Sub


    Private Function deletePrefixe(ByRef edit As String) As Boolean
        For Each i As String In prefixe
            If edit.StartsWith(i) Then
                edit = edit.Substring(i.Length)
                Return True

            End If
        Next

        Return False
    End Function

    Private Function deleteSuffixe(ByRef edit As String) As Boolean
        For Each i As String In suffixe
            If edit.EndsWith(i) Then
                edit = edit.Substring(0, edit.Length - i.Length)
                Return True
            End If
        Next
        Return False
    End Function

    Private Function check(pattrens As String(), word As String)

        For Each i As String In pattrens

            Dim one = i.IndexOf("ف"), two = i.IndexOf("ع"), three = i.IndexOf("ل"), same = word


            Dim arr() As Char = same.ToCharArray()
            arr(one) = "ف"
            arr(two) = "ع"
            arr(three) = "ل"
            same = String.Concat(arr)


            If same = i Then

                Return word(one) + word(two) + word(three)
            End If

        Next
        Return ""
    End Function

    Private Sub firstRemove(ByRef edit As String, ByRef ans As String, ByRef flip As Boolean)
        While ans = "" And flip

            If hashPattren.ContainsKey(edit.Length) Then
                ans = check(hashPattren(edit.Length), edit)
            End If

            If ans = "" Then
                flip = deletePrefixe(edit)
                If flip Then
                    edit.editStartWithPrefixe
                End If
            End If

        End While
    End Sub

    Private Sub secondRemove(ByRef edit As String, ByRef ans As String, ByRef flip As Boolean)
        While ans = "" And flip

            If hashPattren.ContainsKey(edit.Length) Then
                ans = check(hashPattren(edit.Length), edit)
            End If

            If ans = "" Then
                flip = deleteSuffixe(edit)
                If flip Then
                    edit.editEndWithSuffixe
                End If
            End If

        End While
    End Sub

    Private Sub insertInMap(ByRef edit As String, ByRef ans As String, ByRef i As String)


        If edit.Length = 3 And ans = "" Then
            ans = edit
        End If

        If i.Length = 3 And ans = "" Then
            ans = i
        End If

        If ans.Length <> 3 Then
            If (ans.Length <> 3 And i = "الله") Then
            Else
                Exit Sub
            End If
        End If

        If Not newItems.ContainsKey(ans) Then
            newItems.Add(ans, New HashSet(Of String))
        End If
        If Not newItems(ans).Contains(i) Then
            newItems(ans).Add(i)
            createButton(ans, i)
        End If
    End Sub



    Private Sub newForm()
        Me.Hide()
        Me.Size = New Size(1368, 769)
        Me.BackgroundImage = New System.Drawing.Bitmap(Button1.BackgroundImage)
        System.Threading.Thread.Sleep(300)
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



    Private Sub btnSort()
        Dim btn As New Button
        btn.Size = New Size(131, 40)
        btn.Text = "Sort"
        btn.Tag = "Sort"
        btn.Name = "Sort"
        AddHandler btn.Click, AddressOf sort_btn
        btn.Font = New Font(Button1.Font.FontFamily, Button1.Font.Size, Button1.Font.Style)
        btn.Location = New Point(990, 500)
        btn.BackgroundImage = New System.Drawing.Bitmap(Button1.BackgroundImage)
        Me.Controls.Add(btn)
        createBox(listSort)

        listSort.RightToLeft = RightToLeft.Yes
        listSort.Visible = False
    End Sub

    Private Sub createButton(ans As String, i As String)


        Dim btn As New Button
        btn.Size = New Size(131, 40)

        btn.Location = New Point(x, y)
        x += 129
        If x >= 790 Then
            y += 46
            x = 12
        End If
        btn.Text = i
        btn.Tag = ans
        AddHandler btn.MouseHover, AddressOf hover_btn
        AddHandler btn.MouseLeave, AddressOf hover_btn
        panel.Controls.Add(btn)


    End Sub

    Private Sub createBox(ByRef panel As Object)

        panel.Size = New Size(792, 681)

        panel.Location = New Point(2, 12)
        panel.BorderStyle = BorderStyle.Fixed3D
        panel.BackgroundImage = New System.Drawing.Bitmap(Button1.BackgroundImage)
        Me.Controls.Add(panel)
    End Sub

    Private Sub hover_btn(sender As System.Object, e As System.EventArgs) Handles Button1.MouseLeave
        System.Threading.Thread.Sleep(20)
        Dim btn As Button = sender
        Dim newName = btn.Text
        btn.Text = btn.Tag
        btn.Tag = newName

    End Sub



    Private Sub buttonForText()
        Dim btn As New Button
        btn.Size = New Size(131, 40)
        btn.Location = New Point(980, 280)
        btn.BackgroundImage = New System.Drawing.Bitmap(My.Resources._6921087_ai)
        btn.Text = "Add New Text"
        btn.Tag = "Add New Text"
        btn.Name = "btnText"
        btn.AutoSize = True

        AddHandler btn.MouseHover, AddressOf hover_addText
        AddHandler btn.Click, AddressOf click_addText

        Me.Controls.Add(btn)
    End Sub

    Private Sub hover_addText(sender As System.Object, e As System.EventArgs) Handles Button1.MouseHover

        Dim btn As Button = sender
        If btn.Name = "btnText" Then
            TextBox2.Visible = True
        End If
    End Sub

    Private Sub click_addText(sender As System.Object, e As System.EventArgs)


        Dim btn As Button = sender
        If btn.Name = "btnText" Then
            TextBox2.Visible = False
            root(TextBox2.Text.Split(" ").ToList)
            TextBox2.Text = Nothing
        End If

    End Sub

    Private Sub sort_btn(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim btn As Button = sender
        If btn.Name = "Sort" Then
            If flagSort Then
                btn.Text = "Sort"
            Else
                btn.Text = "Un Sort"
            End If
            Me.panel.Visible = Not panel.Visible
            Me.listSort.Visible = Not listSort.Visible
            flagSort = Not flagSort
            printRoot()
        End If
    End Sub

    Private Sub printRoot()


        Dim arr() As String
        If (Not Me.panel.Visible) Then
            listSort.Items.Clear()
            For Each it In newItems
                Dim setWords = String.Join(" , ", it.Value)
                arr = setWords.Split(" , ")
                If arr.Length <= 8 Then
                    listSort.Items.Add(it.Key & " = [ " & setWords & " ]")
                Else
                    Dim firstHalf As String = String.Empty
                    Dim secondHalf As String = String.Empty
                    For i As Integer = 0 To 7
                        firstHalf += arr(i) & " , "
                    Next
                    listSort.Items.Add(it.Key & " = [ " & firstHalf)
                    For i As Integer = 8 To arr.Length - 1
                        secondHalf += arr(i) & " , "
                    Next
                    listSort.Items.Add(secondHalf & " ]")
                End If
            Next
        End If
    End Sub


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            System.Threading.Thread.Sleep(200)
            listSort.Items.Clear()
            For Each i In ignore
                If i.Value Then
                    listSort.Items.Add(i.Key)
                End If
            Next

        Else
            printRoot()
        End If
    End Sub


End Class


'الاسلام المسلمين المسلمات المعامل المعلمين المعلمات المعلم المسلمون 
'من عن لماذا اين كيف هل 
'طفل اطفال الاطفال اطفالكم فأطفالكم اطفالهم والاطفال فاطفالهم وطفل الطفولة والطفلتين طفلتان
