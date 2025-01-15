Imports System.Reflection
Imports System.Runtime.CompilerServices

Module rootEx


    <Extension()>
    Function isDigit(s As String) As Boolean
        For Each i As Char In s
            If (i >= "a" And i <= "z") Or (i >= "A" And i <= "Z") Or (i >= "0" And i <= "9") Then
                Return True
            End If
        Next
        Return False
    End Function

    <Extension()>
    Function isIgnore(s As String, ignore As Dictionary(Of String, Boolean)) As Boolean

        Return If(ignore.ContainsKey(s), True, False)
    End Function

    <Extension()>
    Sub ReplaceNotGoodChar(ByRef s As String)



        Dim arr() As Char = s.ToCharArray
        Dim n As Integer = s.Length - 1

        If arr(0) = "أ" Or arr(0) = "ئ" Then
            arr(0) = "ا"
        End If

        If arr(n) = "ى" Then
            arr(n) = "ي"
        End If
        If arr(n) = "ة" Then
            arr(n) = "ه"
        End If

        s = String.Join("", arr)
        Array.Clear(arr)
    End Sub

    <Extension()>
    Sub editStartWithPrefixe(ByRef S As String)
        Dim arr() As Char = S.ToCharArray
        Dim n As Integer = S.Length - 1

        If arr(0) = "أ" Or arr(0) = "ئ" Then
            arr(0) = "ا"
        End If

        S = String.Join("", arr)
        Array.Clear(arr)
    End Sub


    <Extension()>
    Sub editEndWithSuffixe(ByRef s As String)
        Dim arr() As Char = s.ToCharArray
        Dim n As Integer = s.Length - 1

        If arr(n) = "ئ" Then
            arr(n) = "ا"
        End If
        If arr(n) = "ى" Then
            arr(n) = "ي"
        End If
        If arr(n) = "ة" Or arr(n) = "ت" Then
            arr(n) = "ه"
        End If

        s = String.Join("", arr)
        Array.Clear(arr)
    End Sub

    <Extension()>
    Sub removeSpace(ByRef s As String)
        Dim arr() As Char = s.ToCharArray
        Array.Sort(arr, Function(x As Char, y As Char)
                            Return x = " " Or y = " "
                        End Function)
        Dim i As Integer
        For i = 0 To arr.Length - 1

            If arr(i) >= "ا" And arr(i) <= "ي" Then
                Exit For
            End If
        Next
        'MessageBox.Show(i & s & " " & s.Length)
        s = String.Join("", arr)
        s = s.Substring(i)
        'MessageBox.Show(s & " " & s.Length)
    End Sub

    <Extension()>
    Sub insertIgnoreWords(ignore As Dictionary(Of String, Boolean))
        ignore.Add("اللذين", False)
        ignore.Add("الي", False)
        ignore.Add("اليوم", False)
        ignore.Add("اما", False)
        ignore.Add("امام", False)
        ignore.Add("امس", False)
        ignore.Add("ان", False)
        ignore.Add("او", False)
        ignore.Add("اول", False)
        ignore.Add("اين", False)
        ignore.Add("اي", False)
        ignore.Add("ب", False)
        ignore.Add("بات", False)
        ignore.Add("بان", False)
        ignore.Add("بد", False)
        ignore.Add("بدل", False)
        ignore.Add("بعد", False)
        ignore.Add("بعض", False)
        ignore.Add("بل", False)
        ignore.Add("بيت", False)
        ignore.Add("بين", False)
        ignore.Add("تحت", False)
        ignore.Add("تكون", False)
        ignore.Add("تلك", False)
        ignore.Add("ثم", False)
        ignore.Add("جدا", False)
        ignore.Add("حالي", False)
        ignore.Add("حتى", False)
        ignore.Add("ابا", False)
        ignore.Add("ابو", False)
        ignore.Add("ابي", False)
        ignore.Add("احد", False)
        ignore.Add("اخا", False)
        ignore.Add("اخر", False)
        ignore.Add("اخو", False)
        ignore.Add("اخي", False)
        ignore.Add("اذا", False)
        ignore.Add("الا", False)
        ignore.Add("الان", False)
        ignore.Add("التي", False)
        ignore.Add("الذي", False)
        ignore.Add("الذين", False)
        ignore.Add("اللتان", False)
        ignore.Add("اللتين", False)
        ignore.Add("اللذان", False)
        ignore.Add("حول", False)
        ignore.Add("حيث", False)
        ignore.Add("حين", False)
        ignore.Add("خالل", False)
        ignore.Add("دون", False)
        ignore.Add("ذا", False)
        ignore.Add("ذات", False)
        ignore.Add("ذلك", False)
        ignore.Add("ذو", False)
        ignore.Add("ذي", False)
        ignore.Add("رغم", False)
        ignore.Add("شيء", False)
        ignore.Add("صار", False)
        ignore.Add("صبح", False)
        ignore.Add("صبر", False)
        ignore.Add("ضحى", False)
        ignore.Add("ضد", False)
        ignore.Add("لذلك", False)
        ignore.Add("لعل", False)
        ignore.Add("لكن", False)
        ignore.Add("لم", False)
        ignore.Add("لماذا", False)
        ignore.Add("لن", False)
        ignore.Add("له", False)
        ignore.Add("لو", False)
        ignore.Add("ليت", False)
        ignore.Add("ليس", False)
        ignore.Add("ما", False)
        ignore.Add("ماانفك", False)
        ignore.Add("مابرح", False)
        ignore.Add("ماذا", False)
        ignore.Add("مازال", False)
        ignore.Add("مافتئ", False)
        ignore.Add("مايزال", False)
        ignore.Add("ضمن", False)
        ignore.Add("ظل", False)
        ignore.Add("عل", False)
        ignore.Add("عن", False)
        ignore.Add("عند", False)
        ignore.Add("عين", False)
        ignore.Add("غير", False)
        ignore.Add("ف", False)
        ignore.Add("فقط", False)
        ignore.Add("في", False)
        ignore.Add("فيما", False)
        ignore.Add("قبل", False)
        ignore.Add("قد", False)
        ignore.Add("ك", False)
        ignore.Add("كان", False)
        ignore.Add("كذلك", False)
        ignore.Add("كل", False)
        ignore.Add("كم", False)
        ignore.Add("كون", False)
        ignore.Add("كي", False)
        ignore.Add("كيف", False)
        ignore.Add("ل", False)
        ignore.Add("لا", False)
        ignore.Add("لازال", False)
        ignore.Add("لاسيما", False)
        ignore.Add("لدي", False)
        ignore.Add("لايزال", False)
        ignore.Add("متى", False)
        ignore.Add("مساء", False)
        ignore.Add("مسي", False)
        ignore.Add("مع", False)
        ignore.Add("مما", False)
        ignore.Add("من", False)
        ignore.Add("منذ", False)
        ignore.Add("نحو", False)
        ignore.Add("نفس", False)
        ignore.Add("هؤلاء", False)
        ignore.Add("هذا", False)
        ignore.Add("هذه", False)
        ignore.Add("هل", False)
        ignore.Add("هن", False)
        ignore.Add("هنا", False)
        ignore.Add("هو", False)
        ignore.Add("هي", False)
        ignore.Add("هما", False)
        ignore.Add("هم", False)
        ignore.Add("وسط", False)
        ignore.Add("يكون", False)
        ignore.Add("يلي", False)
        ignore.Add("يمكن", False)
        ignore.Add("يوم", False)

    End Sub

    <Extension()>
    Sub insrtPattrens(hashPattren As Dictionary(Of Integer, String()))
        hashPattren.Add(4, {"فاعل", "فعلى", "فعله", "تفعل", "فعول", "فعيل", "فعال", "مفعل", "افعل", "يفعل"})
        hashPattren.Add(5, {"فاعله", "فاعول", "فعلاء", "فعلان", "فعوله", "فعيله", "فعائل", "فعالي", "فعاله", "فواعل", "افعال", "افعله", "تفعله", "تفعيل", "مفعال", "مفعيل", "مفعول", "مفعله", "مفاعل", "مفتعل", "متفعل", "منفعل"})
        hashPattren.Add(6, {"فعلانه", "تفعيله", "افتعال", "انفعال", "افعالي", "افعالا", "افعلاء", "افعليه", "فاعوله", "فعاليه", "مستفعل", "مفاعله", "مفاعيل", "منفعله", "مفعيله", "مفعوله", "مفتعله", "متفعله", "تفاعيل"})
        hashPattren.Add(7, {"استفعال", "انفعاله", "افتعاله", "افتعالي", "افعاليه", "مستفعله", "مفعوليه", "متفاعله"})
        hashPattren.Add(8, {"استفعالي", "افتعاليه"})
        hashPattren.Add(9, {"استفعاليه"})
    End Sub

    <Extension()>
    Sub insrtPrefixe(ByRef prefixe As String())
        prefixe = {"بال", "فال", "كال", "و", "س", "ف", "وال", "لل", "ال"}
    End Sub

    <Extension()>
    Sub insrtSuffixe(ByRef suffixe As String())
        suffixe = {"هما", "كما", "ات", "يه", "ته", "تي", "ان", "ون", "ين", "هم", "هن", "ها", "نا", "وا", "كم", "كن", "ني", "وني", "تم", "ه", "س", "ي"}
    End Sub


End Module
