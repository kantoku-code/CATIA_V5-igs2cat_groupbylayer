Language = "VBSCRIPT"
'*********************************
'Igs2Cat_GroupByLayer.vbs
'ver 0.02
'by kantoku

'1-CATIA���N��
'2-Igs2Cat_GroupByLayer.vbs�� ����ٸد���
'  Cat-Dll_Env-Path.txt���쐬
'3-Igs2Cat_GroupByLayer.vbs
'  Igs2Cat_Support.CATScript
'  Cat-Dll_Env-Path.txt
'  �𓯈�̫��ނɒu��
'4-Igs2Cat_GroupByLayer.vbs ��Igeş�ق�D&D
'5-Igeş�قƓ���̫��ނ�ڲ԰�������ꂽ��Ԃ�
'  CATPart���o���オ��
'*********************************

'Option Explicit
'*** �ݒ� �ύX���Ȃ��ŉ����� ***
Const ScriptName = "Igs2Cat_GroupByLayer"
Const EnvName = "Cat-Dll_Env-Path.txt"
Const ListFileName = "Igs2Cat_GroupByLayer_IptLst.txt"
Const MacroName = "Igs2Cat_Support.CATScript"
Const EnvKeys = "CAT,DIRENV,ENV,CATTEMP"
'***

Call Main
wscript.Quit 0

'*********************************
Sub Main()
    'D&D
    Dim DDList 'As Variant
    DDList = GetDropList(WScript.Arguments)
    If Not IsArray(DDList) Then Exit Sub '���쐬��Exit

    '���߽�ގ擾
    Dim EnvDic 'As Object
    Set EnvDic = GetDll_Env
    If EnvDic Is Nothing Then Exit Sub

    '�m�F
    Dim DDlistStr 'as String

    Dim Msg 'As String
    Msg = "�ȉ���̧�ق�ϊ����܂��B��낵���ł����H" + vbNewLine + _
      DDList2String(DDList)
    If MsgBox(Msg, vbYesNo, ScriptName) = vbNo Then Exit Sub

    '���߰�ؽč쐬
    Dim ListPath 'As String
    ListPath = Replace(EnvDic("CATTEMP"), Chr(34), "") + "\" + ListFileName
    If IsExists(ListPath) Then
        Msg = "�O��̕ϊ�̧��ؽĂ��c���Ă��܂��B" + vbNewLine + _
          "�㏑�����܂����A��낵���ł��傤���H"
        If MsgBox(Msg, vbYesNo, ScriptName) = vbNo Then Exit Sub
    End If
    Call WriteFile(ListPath, Join(DDList, vbNewLine))

    '�ޯ�Ӱ�ދN��
    Dim MacroPath 'As String
    MacroPath = GetCurrentPath + "\" + MacroName
    Call ExecuteButchMode(EnvDic, MacroPath)

    '��ͻ�߰Ľ��ظĂŏ���
End Sub

' *** ButchMode ***
Private Sub ExecuteButchMode(ByVal Dic, ByVal MacroPath)
    Dim Common 'As String
    Common = Dic("CAT") + " -direnv " + Dic("DIRENV") + _
            " -env " + Dic("ENV") + " -batch  -macro " + _
            Chr(34) + MacroPath + Chr(34)
    Call CreateObject("Wscript.Shell").Exec(Common)
End Sub

' *** Env ***
'�ޯ�Ӱ�ދN���p̧�َ擾
Private Function GetDll_Env() 'As Object
    Dim EnvPath 'As String
    EnvPath = GetCurrentPath+ "\" + EnvName
    If Not IsExists(EnvPath) Then
        Dim Msg 'As String
        Msg = "�ޯ�Ӱ�ދN���ɕK�v��̧�ق�����܂���!" + _
            vbNewLine + "(" + EnvName + ")"
        MsgBox Msg, vbOKOnly, ScriptName
        Set GetDll_Env = Nothing
        Exit Function
    End If
    Dim Txts 'As Variant
    Txts = ReadFile(EnvPath)
    If UBound(Txts) < 3 Then Exit Function

    Dim Dic 'As Object
    Set Dic = CreateObject("Scripting.Dictionary")

    Dim I 'As Long
    Dim KeyValue 'As Variant
    For I = 0 To UBound(Txts)
        KeyValue = GetKeyValue(Txts(I))
        If Not UBound(KeyValue) = 1 Then Exit Function
        Dic.Add KeyValue(0), KeyValue(1)
    Next
    If Not CheckEnv(Dic) Then
        Set GetDll_Env = Nothing
        Exit Function
    End If
    Set GetDll_Env = Dic
End Function

'��������
Private Function CheckEnv(ByVal Dic) 'As Boolean
    Dim I 'As Long
    Dim AryEnvKeys 'As Variant
    AryEnvKeys = Split(EnvKeys, ",")
    For I = 0 To UBound(AryEnvKeys)
        If Not Dic.Exists(AryEnvKeys(I)) Then
            Dim Msg 'As String
            Msg = "�ޯ�Ӱ�ދN���ɕK�v��̧�ٓ��̐ݒ肪����܂���!" + _
                vbNewLine + "(" + AryEnvKeys(I) + ")"
            MsgBox Msg, vbOKOnly, ScriptName
            CheckEnv = False
            Exit Function
        End If
    Next
    CheckEnv = True
End Function

'�N���pKeyValue
'Return: 0-Key 1-Value
Private Function GetKeyValue(ByVal Txt) 'As Variant
    Dim Equal 'As Variant
    Equal = Split(Txt, "=")
    If Not UBound(Equal) = 1 Then Exit Function

    Dim Spece 'As Variant
    Spece = Split(Equal(0), " ")
    If Not UBound(Spece) = 1 Then Exit Function

    Dim KeyValue(1) 'As String
    KeyValue(0) = Spece(1)
    KeyValue(1) = Equal(1)
    GetKeyValue = KeyValue
End Function


' *** D&D ***
'��ۯ�ߏ���
Private Function GetDropList(ByVal Args) 'As Variant
    Dim ArgsCount 'As Long
    ArgsCount = Args.Count
    If ArgsCount < 1 Then
        Call GetEnvMain()
        Exit Function
    End If

    Dim I 'As Long
    Dim IgsList() 'As Variant
    ReDim IgsList(ArgsCount)
    Dim IgsCount 'As Long
    IgsCount = -1
    Dim Path 'As Variant
    Dim ArgsPath 'As String

    'Continue��Goto�g�������������
    For I = 1 To ArgsCount
        ArgsPath = Args(I - 1)
        If IsExists(ArgsPath) Then
            Path = SplitPathName(ArgsPath)
            Path = AskIges(Path)
            If IsIgsFile(Path(2)) Then
                IgsCount = IgsCount + 1
                IgsList(IgsCount) = JoinPathName(Path)
            End If
        End If
    Next

    If IgsCount < 0 Then
        Msg = "�ϊ��\��̧�ق�����܂���!"
        MsgBox Msg, vbOKOnly, ScriptName
        Exit Function
    End If

    ReDim Preserve IgsList(IgsCount)
    GetDropList = IgsList
End Function

'�g���q .iges �ւ̑Ή�
Private Function AskIges(ByVal Path) 'As Variant
    If UCase(Path(2)) = "IGES" Then
        Dim Msg 'As String
        Msg = JoinPathName(Path) + vbNewLine + _
            "���A�g���q igs �ɕύX���ĕϊ����܂���?"
        If MsgBox(Msg, vbYesNo, ScriptName) = vbYes Then
            Path = ChengeExt2Igs(Path)
        End If
    End If
    AskIges = Path
End Function

'Iges�����@�g���q�̂� iif()�g������
Private Function IsIgsFile(ByVal Ext) 'As Boolean
    IsIgsFile = False
    If UCase(Ext) = "IGS" Then IsIgsFile = True
End Function

'ؽĂ�̧��Ҳ�̂ݎ擾
Private Function DDList2String(ByVal DDlist) 'As Boolean
    Dim Ts,ToStr,i
    ToStr= ""
    For i = 0 to UBound(DDlist)
        Ts = SplitPathName(DDlist(i))
        ToStr = ToStr+Ts(1) + "." + Ts(2) + vbNewLine
    Next
    DDList2String = ToStr
End Function


' *** IO ***
'FileSystemObject
Private Function GetFSO() 'As Object
    Set GetFSO = CreateObject("Scripting.FileSystemObject")
End Function

'�߽/̧�ٖ�/�g���q ����
'Return: 0-Path 1-BaseName 2-Extension
Private Function SplitPathName(ByVal FullPath) 'As Variant
    Dim Path(2) 'As String
    With GetFSO
        Path(0) = .GetParentFolderName(FullPath)
        Path(1) = .GetBaseName(FullPath)
        Path(2) = .GetExtensionName(FullPath)
    End With
    SplitPathName = Path
End Function

'�߽/̧�ٖ�/�g���q �A��
Private Function JoinPathName(ByVal Path) 'As String
    If Not IsArray(Path) Then Stop '���Ή�
    If Not UBound(Path) = 2 Then Stop '���Ή�
    JoinPathName = Path(0) + "\" + Path(1) + "." + Path(2)
End Function

'̧�ق̗L��
Private Function IsExists(ByVal Path) 'As Boolean
    IsExists = GetFSO.FileExists(Path)
End Function

'�g���q�ύX �ǂݍ��ݐ�p����
Private Function ChengeExt2Igs(ByVal Path) 'As Variant
    Dim NewName 'As String
    NewName = Path(1) + ".igs"
    Dim NewPath 'As String
    NewPath = Path(0) + "\" + NewName
    If IsExists(NewPath) Then
        Dim Msg 'As String
        Msg = NewPath + vbNewLine + "�����ɗL�邽�߁A�ύX�ł��܂���ł���"
        MsgBox Msg, vbOKOnly, ScriptName
    Else
        GetFSO.GetFile(JoinPathName(Path)).Name = NewName
        Path(2) = "igs"
    End If
    ChengeExt2Igs = Path
End Function

'̧�ٓǂݍ���
Private Function ReadFile(ByVal Path) 'As Variant
    With GetFSO.GetFile(Path).OpenAsTextStream
        ReadFile = Split(.ReadAll, vbNewLine)
        .Close
    End With
End Function

'̧�ُ�������
Private Sub WriteFile(ByVal Path, ByVal Txt)
    With GetFSO.OpenTextFile(Path, 2, True)
        .Write Txt
        .Close
    End With
End Sub

'�������߽
Private Function GetCurrentPath() 'As String
    GetCurrentPath = GetFSO.getParentFolderName(WScript.ScriptFullName)
End Function

' *** ���擾 ***
Private Sub GetEnvMain()
    '���߽���擾����CATIA�̎擾
    Dim Cat 'As Application
    Set Cat = GetCatia()
    If Cat Is Nothing Then Exit Sub

    'catia�̎��ş���߽�擾
    Dim CatPath ' As String
    CatPath = Cat.SystemService.Environ("CATDLLPath")

    '��̧���߽�擾
    Dim EnvironmentPath ' As Variant
    EnvironmentPath = SplitPathName(Cat.SystemService.Environ("CATEnvName"))

    'TEMP̫����߽�擾
    Dim TempPath ' As Variant
    TempPath = Cat.SystemService.Environ("CATTemp")

    '�o�͕���
    Dim ExpTxt ' As String
    ExpTxt = "Set CAT=" + Chr(34) + CatPath + "\CNEXT.exe" + Chr(34) + vbNewLine + _
             "Set DIRENV=" + Chr(34) + EnvironmentPath(0) + Chr(34) + vbNewLine + _
             "Set ENV=" + Chr(34) + EnvironmentPath(1) + Chr(34) + vbNewLine + _
             "Set CATTEMP=" + Chr(34) + TempPath + Chr(34)

    '�ۑ�
    Dim ExpPath 'As String
    ExpPath = GetCurrentPath + "\" + EnvName
    If IsExists(ExpPath) Then
        Dim Msg 'As String
        Msg = "�u" + EnvName + "�v�����݂��܂��B�㏑�����܂���?(������-��ݾ�)"
        If MsgBox(Msg, vbYesNo, ScriptName) = vbNo Then Exit Sub
    End If
    Call WriteFile(ExpPath, ExpTxt)

    '�I��
    MsgBox ExpPath + vbNewLine + "���쐬���܂���", vbOKOnly, ScriptName
End Sub

'�N������catia�̎擾
Private Function GetCatia() 'As Application
    On Error Resume Next
        Set GetCatia = GetObject(, "CATIA.Application")
        If GetCatia Is Nothing Then
            MsgBox "CATIA V5 ���N�����Ă�������", vbOKOnly, ScriptName
            Err.Clear
            wscript.Quit 0
        End If
    On Error GoTo 0
End Function
