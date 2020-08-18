Imports Shell32
Imports System.IO

Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization
Imports System.Runtime.CompilerServices

Imports System.Runtime.InteropServices

Module Module1
    Sub Main(args As String())
    '
    ' 
    ' ConvertHeosPlaylistToM3u
    '
    ' HEOSコマンドで取得したJSON型式のプレイリスト(Playlist)を元に
    ' ファイルパス情報を追加したm3u型式のプレイリストを生成する
    '
    '
    ' [第1引数]：HEOSコマンドで取得したJSON型式のプレイリストファイルパス
    '  Dim PlayListFilePath As String = "D:\tmp\Playlist 1.txt"
    '
    ' [第2引数]：メディアファイルのROOTフォルダパス
    '  Dim targetFolder As String = "F:\DATA\"
    '
    ' [第3引数]：メディアファイルのMP3Tag情報を保存するキャッシュファイルパス
    '  Dim MP3TagInfoFilePath As String = "F:\tmp\MP3FileTags.txt"
    '
    '  キャッシュファイルが存在しなければメディアファイルのROOTフォルダパスから
    '  全てのメディアファイルの情報を元に作成される
    '
    '
    ' ■処理概要
    ' 1. 指定したフォルダ以下に含まれるメディアファイルのMP3Tag情報と
    '    そのファイルへのパス情報をキャッシュファイルに集める。
    '    キャッシュファイルが存在すれば、この処理はスキップする。
    '
    ' 2. HEOSプレイリストからm3u形式のプレイリストを生成する
    '    プレイリストとMP3のTag情報(キャッシュファイル)と比較し、
    '    一致したらm3u形式のプレイリストに追加する
    '
    '    ※m3u型式ファイルの出力フォルダは、第2引数で指定したROOTフォルダパス直下に
    '      #Playlists# というフォルダを作成する。
    '    ※m3u型式ファイルのファイル名は、第1引数で指定したプレイリストファイルの
    '      拡張子をm3uに置き換えたもの。
    '
    '
    ' ■ビルド環境
    '  VisualStudio VB.NET Windowsコンソールプロジェクト
    ' ・参照の追加(COM)：Microsoft Shell Controls And Automation
    ' ・参照の追加(アセンブリ拡張)：Json.NET is a popular high-performance JSON framework for .NET
    ' 
    ' ■修正履歴
    ' 2020.08.11 コマンドライン引数起動バージョン
    ' 2020.08.11 終了時 Console.ReadLine()でなく、終了メッセージ表示に変更
    ' 2020.08.12 参照サイトを明記
    '
    '
    ' Copy Right (C) Hiroyasu Watanabe 2020.0.13
    '


    '
    '1. 指定したフォルダ以下に含まれるメディアファイルのMP3Tag情報とそのファイルへのパス情報を集める
    '
    If args.Length = 0 Then
            Console.WriteLine("コマンドライン引数はありません。")
        Else

#If DEBUG Then

            Dim arg As String
            For Each arg In args
                Console.WriteLine(arg)
            Next
            Console.WriteLine("")

#End If

            Console.WriteLine("第1引数：" & args(0))
            Console.WriteLine("第2引数：" & args(1))
            Console.WriteLine("第3引数：" & args(2))

        End If

        'メディアファイルのROOTフォルダパス
        Dim targetFolder As String = args(1)

        'メディアファイルのMP3Tag情報を保存するファイル
        '存在しなければメディアファイルのROOTフォルダパスから
        '全てのメディアファイルの情報を元に作成される
        Dim MP3TagInfoFilePath As String = args(2)

        'm3u出力フォルダ(メディアファイルから相対パスで指定できる場所を指定する)
        Dim m3uOutputFolder As String = targetFolder & "#Playlists#\"


        If System.IO.File.Exists(MP3TagInfoFilePath) Then
            'MP3TagInfoFileキャッシュファイルが存在したら上書きせず使用する
            GoTo create_M3U_only
        End If


        'MP3TagInfoFileファイルがない場合は作成する
        'メディアファイルパス情報をMP3TagInfoFileファイルへ追加
        Dim filePattern As String = "*.mp3"
        AddMediaFilePathToMP3TagInfoFile(targetFolder, filePattern, MP3TagInfoFilePath)
        'メディアファイルパス情報をMP3TagInfoFileファイルへ追加
        filePattern = "*.m4a"
        AddMediaFilePathToMP3TagInfoFile(targetFolder, filePattern, MP3TagInfoFilePath)
        'メディアファイルパス情報をMP3TagInfoFileファイルへ追加
        filePattern = "*.flac"
        AddMediaFilePathToMP3TagInfoFile(targetFolder, filePattern, MP3TagInfoFilePath)


create_M3U_only:

        '2. HEOSプレイリストからm3u形式のプレイリストを生成する
        '   プレイリストとMP3のTag情報と比較し、一致したらm3u形式のプレイリストに追加する

        Dim PlayListFilePath As String = args(0)
        ConvertJsonPlaylistToM3U(PlayListFilePath, MP3TagInfoFilePath, m3uOutputFolder)

        'Console.ReadLine()
        Console.WriteLine("正常に終了しました。")

    End Sub

    <DllImport("shlwapi.dll", CharSet:=CharSet.Auto)>
    Private Function PathRelativePathTo(
    <Out> pszPath As StringBuilder,
    <[In]> pszFrom As String,
    <[In]> dwAttrFrom As System.IO.FileAttributes,
    <[In]> pszTo As String,
    <[In]> dwAttrTo As System.IO.FileAttributes) As Boolean
    End Function
    ''' ---------------------------------------------------------------------------------------
    ''' ※参照記事：DOBON.NET
    ''' https://dobon.net/vb/dotnet/file/getabsolutepath.html
    ''' <summary>
    ''' 絶対パスから相対パスを取得します。
    ''' </summary>
    ''' <param name="basePath">基準とするフォルダのパス。</param>
    ''' <param name="absolutePath">相対パス。</param>
    ''' <returns>絶対パス。</returns>
    Public Function GetRelativePath(basePath As String,
        absolutePath As String) As String
        Dim sb As New StringBuilder(260)
        Dim res As Boolean = PathRelativePathTo(
        sb, basePath, System.IO.FileAttributes.Directory,
        absolutePath, System.IO.FileAttributes.Normal)
        If Not res Then
            Throw New Exception("相対パスの取得に失敗しました。")
        End If
        Return sb.ToString()
    End Function

    ''' ---------------------------------------------------------------------------------------
    ''' ※参照記事：フォルダ以下のファイルを最下層まで検索または取得する
    ''' http://jeanne.wankuma.com/tips/vb.net/directory/getfilesmostdeep.html
    ''' <summary>
    '''     指定した検索パターンに一致するファイルを最下層まで検索しすべて返します。</summary>
    ''' <param name="stRootPath">
    '''     検索を開始する最上層のディレクトリへのパス。</param>
    ''' <param name="stPattern">
    '''     パス内のファイル名と対応させる検索文字列。</param>
    ''' <returns>
    '''     検索パターンに一致したすべてのファイルパス。</returns>
    ''' ---------------------------------------------------------------------------------------
    Public Function GetFilesMostDeep(ByVal stRootPath As String, ByVal stPattern As String) As String()
        Dim hStringCollection As New System.Collections.Specialized.StringCollection()

        ' このディレクトリ内のすべてのファイルを検索する
        For Each stFilePath As String In System.IO.Directory.GetFiles(stRootPath, stPattern)
            hStringCollection.Add(stFilePath)
        Next stFilePath

        ' このディレクトリ内のすべてのサブディレクトリを検索する (再帰)
        For Each stDirPath As String In System.IO.Directory.GetDirectories(stRootPath)
            Dim stFilePathes As String() = GetFilesMostDeep(stDirPath, stPattern)

            ' 条件に合致したファイルがあった場合は、ArrayList に加える
            If Not stFilePathes Is Nothing Then
                hStringCollection.AddRange(stFilePathes)
            End If
        Next stDirPath

        ' StringCollection を 1 次元の String 配列にして返す
        Dim stReturns As String() = New String(hStringCollection.Count - 1) {}
        hStringCollection.CopyTo(stReturns, 0)

        Return stReturns
    End Function

    Public Sub ConvertJsonPlaylistToM3U(jsonFilePath As String, MP3TagInfoFilePath As String, m3uOutputFolder As String)
        '
        ' 指定したHEOSプレイリストからm3u形式のプレイリストを生成する
        ' プレイリストとMP3のTag情報(キャッシュファイル)と比較し、
        ' 一致したらm3u形式のプレイリストに追加する
        '
        '
        ' [第1引数]：HEOSコマンドで取得したJSON型式のプレイリストファイルパス
        '  Dim PlayListFilePath As String = "D:\tmp\Playlist 1.txt"
        ' 
        ' [第2引数]：メディアファイルのMP3Tag情報の保存されたキャッシュファイルパス
        '  Dim MP3TagInfoFilePath As String = "F:\tmp\MP3FileTags.txt"
        '  ファイルフォーマットはJSON型式
        '
        ' [第3引数]：m3u形式のプレイリストの出力フォルダパス
        '  Dim m3uOutputFolder As String = "F:\DATA\#Playlists#\"
        '
        '  ※m3u型式ファイルのファイル名は、jsonFilePath の拡張子をm3uに置き換えたもの
        '  ※メディアファイルが見つからない曲のログファイル名は"jsonFilePath の拡張子無しファイル名"_NotFoundMediaList.txt
        '
        '■処理概要
        ' ・jsonFilePathで指定されたHEOSの1つのプレイリストは対応する1つのm3uファイルに出力される
        ' ・jsonFilePathで指定されたHEOSの1つのプレイリストは複数のJsonObjectから構成される
        ' ・JsonObjectは各1行
        ' ・1行に分解してからパースしてm3uファイルに追加する

        Dim enc As Encoding = Encoding.UTF8

        Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(jsonFilePath)
        Dim M3UFilePath As String = m3uOutputFolder & fileName & ".m3u"
        Dim NotFoundMediaFilePath As String = m3uOutputFolder & fileName & "_NotFoundMediaList.txt"

        'フォルダ生成
        If Not System.IO.Directory.Exists(m3uOutputFolder) Then
            System.IO.Directory.CreateDirectory(m3uOutputFolder)
        End If

        'm3uファイルクリア、ヘッダ出力
        Dim mediaNotFoundCount As Long = 0
        Using sw As New System.IO.StreamWriter(M3UFilePath)
            sw.WriteLine("#EXTM3U")
        End Using

        'メディアファイルがない曲のリスト出力ファイルクリア
        Using sw2 As New System.IO.StreamWriter(NotFoundMediaFilePath)
        End Using

        Dim jsonStr As String

        'ファイルからJson文字列を読み込む
        Using sr As New System.IO.StreamReader(jsonFilePath, enc)

            While sr.Peek() >= 0

                jsonStr = sr.ReadLine()
                jsonStr = jsonStr.Trim()

                If jsonStr.StartsWith("{") And jsonStr.EndsWith("}") Then

                    'Json文字列をJson形式データに復元する
                    Dim jsonObj As Object = JsonConvert.DeserializeObject(jsonStr)
                    AddJsonObjectToM3U(jsonObj, MP3TagInfoFilePath, M3UFilePath, NotFoundMediaFilePath, mediaNotFoundCount)

                End If

            End While

        End Using

        'Console.ReadLine()

    End Sub

    Private Sub AddJsonObjectToM3U(jsonObj As Object, MP3TagInfoFilePath As String, M3UFilePath As String, NotFoundMediaFilePath As String, ByRef mediaNotFoundCount As Long)
        '
        ' 指定したJsonObjectに含まれる全曲について一致した曲のメディアファイルパスをm3uファイルに追加する
        '
        ' JsonObjectに含まれる曲を1曲ずつ取り出しMP3のTag情報(キャッシュファイル)と比較し、
        ' 一致したらそのファイルパスをm3u形式のプレイリストに追加する
        '
        '
        ' [第1引数]：複数の「曲名、アーティスト名、アルバム名」情報を含んだJsonObject
        ' 
        ' [第2引数]：メディアファイルのMP3Tag情報の保存されたキャッシュファイルパス
        '  Dim MP3TagInfoFilePath As String = "F:\tmp\MP3FileTags.txt"
        '  ファイルフォーマットはJSON型式
        '
        ' [第3引数]：m3u形式のプレイリストの出力ファイルパス
        '  Dim M3UFilePath As String = "F:\DATA\#Playlists#\Playlist 1.m3u"
        '
        ' [第4引数]：メディアファイルが見つからない曲のログファイルパス
        '  Dim NotFoundMediaFilePath As String = "F:\DATA\#Playlists#\Playlist 1_NotFoundMediaList.txt"
        '
        ' [第5引数]：1つのm3u形式のプレイリストでメディアファイルが見つからない曲の合計数
        '  Dim mediaNotFoundCount As Long
        '
        '
        '■処理概要
        ' ・JsonObjectは複数の曲から構成される
        ' ・JsonObjectから曲「曲名、アーティスト名、アルバム名」を取り出し、キャッシュファイルに含まれているか調べる
        ' ・一致した曲はメディアファイルパスをm3uファイルに追加する
        ' ・但し、メディアファイルパスは相対パスに変換して追加する
        '   相対パスはm3uファイルのあるフォルダ("メディアファイルのROOTフォルダパス"\#Playlists#)を原点とする

        'Console.WriteLine("payload={0}", jsonObj("payload"))

        Dim filePath As String
        Dim fileName As String
        Dim targetFolder As String
        Dim m3uBaseFolder As String = System.IO.Path.GetDirectoryName(M3UFilePath) & "\"
        Dim relativePath As String

        Dim comment As String

        If jsonObj("payload") IsNot Nothing Then

            Using sw As New System.IO.StreamWriter(M3UFilePath, True)

                Using sw2 As New System.IO.StreamWriter(NotFoundMediaFilePath, True)

                    For Each item In jsonObj("payload")

                        If item("type") = "song" Then

                            Console.WriteLine("タイトル={0} アーティスト={1} アルバム={2}", item("name"), item("artist"), item("album"))

                            'プレイリスト情報と一致するMP3タグを探してメディアファイルのパスを取得する
                            filePath = SearchMP3TagInfoFile(MP3TagInfoFilePath, item("name"), item("artist"), item("album"), False)

                            '一致しなかったら異なるアルバムも探してみる
                            'If filePath = "" Then
                            '   filePath = SearchMP3TagInfoFile(MP3TagInfoFilePath, item("name"), item("artist"), item("album"), True)
                            'End If

                            If filePath <> "" Then
                                Console.WriteLine(filePath)

                                'm3uファイル出力
                                fileName = System.IO.Path.GetFileName(filePath)
                                sw.WriteLine("#EXTINF:0, " & fileName)

                                targetFolder = System.IO.Path.GetDirectoryName(filePath) & "\"

                                relativePath = GetRelativePath(m3uBaseFolder, targetFolder)

                                filePath = filePath.Replace(targetFolder, relativePath)

                                '相対パスで書き込み
                                sw.WriteLine(filePath)

                                '改行
                                sw.WriteLine("")

                            Else

                                mediaNotFoundCount += 1

                                Console.WriteLine("★★★★: " & mediaNotFoundCount)

                                'm3uファイル出力
                                sw.WriteLine("#EXTINF:0, {0} - {1}", item("artist"), item("name"))

                                '見つからなかったコメント出力
                                comment = "#COMMENT NOT_FOUND_MEDIA_FILE:" & mediaNotFoundCount & ","
                                sw.WriteLine("{0} タイトル={1} アーティスト={2} アルバム={3}", comment, item("name"), item("artist"), item("album"))

                                '改行
                                sw.WriteLine("")

                                '見つからなかった曲情報をエラーログに出力
                                sw2.WriteLine("{0} タイトル={1} アーティスト={2} アルバム={3}", comment, item("name"), item("artist"), item("album"))

                            End If

                        End If

                    Next

                End Using

            End Using

        End If


    End Sub

    Private Function SearchMP3TagInfoFile(MP3TagInfoFilePath As String, title As String, artist As String, album As String, bSearchOtherAlbum As Boolean) As String
        '
        ' 指定した「曲名、アーティスト名、アルバム名」をキーにMP3のTag情報(キャッシュファイル)を検索し
        ' 全て一致したら曲のメディアファイルパスを返す
        '
        ' 
        ' [第1引数]：メディアファイルのMP3Tag情報の保存されたキャッシュファイルパス
        '  Dim MP3TagInfoFilePath As String = "F:\tmp\MP3FileTags.txt"
        '  ファイルフォーマットはJSON型式
        '
        ' [第2引数]：検索したい曲名
        ' [第3引数]：検索したいアーティスト名
        ' [第4引数]：検索したいアルバム名
        '
        ' [第5引数]：「アルバム名」検索条件オプション
        '  ・Trueであれば「アルバム名」が一致しなくても「曲名」と「アーティスト名」のみ一致するまで検索する
        '  ・Falseであれば「アルバム名」と「曲名」と「アーティスト名」の全てが一致するまで検索する
        '  Dim bSearchOtherAlbum As Boolean
        '
        ' [リターン値]：検索条件が一致したメディアファイルパス
        '
        '
        '■処理概要
        ' ・メディアファイルのMP3Tag情報の保存されたキャッシュファイルは、複数のJsonObjectから構成される
        ' ・JsonObjectは複数の曲から構成される
        ' ・JsonObjectから曲「曲名、アーティスト名、アルバム名」を取り出し、
        '   引数で指定された曲「曲名、アーティスト名、アルバム名」と比較する
        ' ・比較する際は、第5引数に指定した検索オプションに従う
        ' ・比較する際は、アルファベットの場合UPPER CASE(大文字)に変換してから比較する
        ' ・一致したら"media_path"プロパティの値をファイルパスとして返す
        '

        Dim enc As Encoding = Encoding.UTF8
        Dim jsonStr As String

        Dim mediaFilePath As String = ""
        Dim tagArtist As String
        Dim tagTitle As String
        Dim tagAlbum As String

        'ファイルからJson文字列を読み込む
        Using sr As New System.IO.StreamReader(MP3TagInfoFilePath, enc)

            While sr.Peek() >= 0

                jsonStr = sr.ReadLine()
                jsonStr = jsonStr.Trim()

                If jsonStr.StartsWith("{") And jsonStr.EndsWith("}") Then

                    'Json文字列をJson形式データに復元する
                    Dim jsonObj As Object = JsonConvert.DeserializeObject(jsonStr)

                    If jsonObj("payload") IsNot Nothing Then

                        '配列の場合
                        For Each item In jsonObj("payload")

                            If item("type") = "song" Then

                                tagArtist = item("artist")
                                tagTitle = item("name")
                                tagAlbum = item("album")

                                'Console.WriteLine("タイトル={0} アーティスト={1} アルバム={2}", item("name"), item("artist"), item("album"))

                                artist = artist.ToUpper()
                                title = title.ToUpper()
                                album = album.ToUpper()

                                If artist = tagArtist.ToUpper() And
                                   title = tagTitle.ToUpper() And
                                   (album = tagAlbum.ToUpper() Or bSearchOtherAlbum) Then

                                    mediaFilePath = item("media_path")

                                    Exit For

                                End If

                            End If

                        Next

                    End If

                End If


                If mediaFilePath <> "" Then
                    Exit While
                End If

            End While

        End Using


        Return mediaFilePath

    End Function
    Private Sub DeleteMP3TagInfoFile(MP3TagInfoFilePath As String)

        If System.IO.File.Exists(MP3TagInfoFilePath) Then
            System.IO.File.Delete(MP3TagInfoFilePath)
        End If

    End Sub
    Private Sub AddMediaFilePathToMP3TagInfoFile(targetFolder As String, filePattern As String, MP3TagInfoFilePath As String)
        '
        ' 指定したフォルダから、指定したファイルパターンを含むファイルを最下層まで検索し
        ' キャッシュファイルに取得する
        '
        '
        ' [第1引数]：メディアファイルのROOTフォルダパス
        '  Dim targetFolder As String = "F:\DATA\"
        '  mp3やm4aなどを検索する最上位階層フォルダ
        ' 
        ' [第2引数]：検索するファイルパターン。ワイルドカード*使用可。
        '  Dim filePattern As String = "*.mp3"
        '
        ' [第3引数]：メディアファイルのMP3Tag情報を保存するキャッシュファイルパス
        '  Dim MP3TagInfoFilePath As String = "F:\tmp\MP3FileTags.txt"
        '  ファイルフォーマットはJSON型式
        '
        '
        Dim stFilePathes As String() = GetFilesMostDeep(targetFolder, filePattern)
        Dim stPrompt As String = String.Empty

        Dim dir As String
        Dim shell As New Shell()
        Dim f As Folder
        Dim file As String
        Dim item As FolderItem

        Dim returnPath As String = ""

        Dim tagArtist As String
        Dim tagTitle As String
        Dim tagAlbum As String

        Dim jsonStr As String = ""

        Dim enc As Encoding = Encoding.UTF8


        'フォルダ生成
        Dim MP3TagInfoFolder As String = System.IO.Path.GetDirectoryName(MP3TagInfoFilePath)
        If Not System.IO.Directory.Exists(MP3TagInfoFolder) Then
            System.IO.Directory.CreateDirectory(MP3TagInfoFolder)
        End If

        'ファイルへJson文字列を書き込む
        'Append = true
        Using sw As New System.IO.StreamWriter(MP3TagInfoFilePath, True)

            ' 取得したファイル名を列挙する
            For Each stFilePath As String In stFilePathes

                If stFilePath Is stFilePathes.First Then

                    jsonStr = "{""payload"": ["
                    sw.Write(jsonStr)

                End If

                dir = Path.GetDirectoryName(stFilePath)
                f = shell.NameSpace(dir)
                file = Path.GetFileName(stFilePath)
                item = f.ParseName(file)

                Console.WriteLine("---- " & stFilePath)

                'Console.WriteLine(f.GetDetailsOf(item, 13))  ' アーティスト
                'Console.WriteLine(f.GetDetailsOf(item, 21))  ' タイトル
                'Console.WriteLine(f.GetDetailsOf(item, 14))  ' アルバムのタイトル

                'Console.WriteLine(f.GetDetailsOf(item, 237)) ' アルバムのアーティスト
                'Console.WriteLine(f.GetDetailsOf(item, 26))  ' トラック番号
                'Console.WriteLine(f.GetDetailsOf(item, 15))  ' 年
                'Console.WriteLine(f.GetDetailsOf(item, 16))  ' ジャンル
                'Console.WriteLine(f.GetDetailsOf(item, 24))  ' コメント

                'ここで時間がかかっている
                tagArtist = f.GetDetailsOf(item, 13)
                tagTitle = f.GetDetailsOf(item, 21)
                tagAlbum = f.GetDetailsOf(item, 14)

                tagArtist = EscapeJSON(tagArtist)
                tagTitle = EscapeJSON(tagTitle)
                tagAlbum = EscapeJSON(tagAlbum)

                Dim stFilePathEscape = stFilePath.Replace("\", "\\")

                jsonStr = "{""container"": ""no"", ""type"": ""song"", ""artist"": """ & tagArtist _
                        & """, ""name"": """ & tagTitle & """, ""album"": """ & tagAlbum & """, ""media_path"": """ & stFilePathEscape & """}"

                If stFilePath Is stFilePathes.Last Then

                    'do something with your last item'
                    jsonStr &= "]}"
                    sw.WriteLine(jsonStr)

                Else

                    'do something with your not last item'
                    jsonStr &= ", "
                    sw.Write(jsonStr)

                End If

            Next stFilePath

        End Using

    End Sub

    Private Function EscapeJSON(str As String)
        '
        'JSON変換前のエスケープ処理関数
        '
        str = str.Replace("\", "\\")
        str = str.Replace("/", "\/")
        str = str.Replace("""", "\""")

        Return str

    End Function

    Sub TestMain()
        '
        ' テスト関数
        '
        ' 複数のプレイリストを生成する
        '
        ' 1. 指定したフォルダ以下に含まれるメディアファイルのMP3Tag情報とそのファイルへのパス情報を集める
        ' 2. プレイリスト毎にm3u形式のプレイリストを生成する
        '   プレイリストとMP3のTag情報と比較し、一致したらm3u形式のプレイリストに追加する
        '
        ' Copyright by nandemo company 2020.07.28
        '

        '1. 指定したフォルダ以下に含まれるメディアファイルのMP3Tag情報とそのファイルへのパス情報を集める

        'メディアファイルのROOTフォルダパス
        Dim targetFolder As String = "F:\DATA\"

        'メディアファイルのMP3Tag情報を保存するファイル
        Dim MP3TagInfoFilePath As String = "F:\tmp\MP3FileTags.txt"

        'm3u出力フォルダ(メディアファイルから相対パスで指定できる場所を指定する)
        Dim m3uOutputFolder As String = targetFolder & "#Playlists#\"


        'MP3TagInfoFileファイルは更新する場合はコメントアウト！！
        GoTo create_M3U_only


        'MP3TagInfoFileファイルをクリアするため削除
        DeleteMP3TagInfoFile(MP3TagInfoFilePath)

        'メディアファイルパス情報をMP3TagInfoFileファイルへ追加
        Dim filePattern As String = "*.mp3"
        AddMediaFilePathToMP3TagInfoFile(targetFolder, filePattern, MP3TagInfoFilePath)
        'メディアファイルパス情報をMP3TagInfoFileファイルへ追加
        filePattern = "*.m4a"
        AddMediaFilePathToMP3TagInfoFile(targetFolder, filePattern, MP3TagInfoFilePath)
        'メディアファイルパス情報をMP3TagInfoFileファイルへ追加
        filePattern = "*.flac"
        AddMediaFilePathToMP3TagInfoFile(targetFolder, filePattern, MP3TagInfoFilePath)


create_M3U_only:

        '2. プレイリスト毎にm3u形式のプレイリストを生成する
        '   プレイリストとMP3のTag情報と比較し、一致したらm3u形式のプレイリストに追加する

        Dim PlayListFilePath As String = "D:\tmp\Playlist #1.txt"
        ConvertJsonPlaylistToM3U(PlayListFilePath, MP3TagInfoFilePath, m3uOutputFolder)

        PlayListFilePath = "D:\tmp\Playlist DISCO1.txt"
        ConvertJsonPlaylistToM3U(PlayListFilePath, MP3TagInfoFilePath, m3uOutputFolder)

        PlayListFilePath = "D:\tmp\Playlist DISCO2.txt"
        ConvertJsonPlaylistToM3U(PlayListFilePath, MP3TagInfoFilePath, m3uOutputFolder)

        PlayListFilePath = "D:\tmp\Sonny Criss.txt"
        ConvertJsonPlaylistToM3U(PlayListFilePath, MP3TagInfoFilePath, m3uOutputFolder)

        PlayListFilePath = "D:\tmp\Bach.txt"
        ConvertJsonPlaylistToM3U(PlayListFilePath, MP3TagInfoFilePath, m3uOutputFolder)

        PlayListFilePath = "D:\tmp\Chopin.txt"
        ConvertJsonPlaylistToM3U(PlayListFilePath, MP3TagInfoFilePath, m3uOutputFolder)

        PlayListFilePath = "D:\tmp\Mozart.txt"
        ConvertJsonPlaylistToM3U(PlayListFilePath, MP3TagInfoFilePath, m3uOutputFolder)

        PlayListFilePath = "D:\tmp\SCANDAL Selection.txt"
        ConvertJsonPlaylistToM3U(PlayListFilePath, MP3TagInfoFilePath, m3uOutputFolder)

        PlayListFilePath = "D:\tmp\SpecialThanks.txt"
        ConvertJsonPlaylistToM3U(PlayListFilePath, MP3TagInfoFilePath, m3uOutputFolder)

        Console.ReadLine()

    End Sub

End Module

