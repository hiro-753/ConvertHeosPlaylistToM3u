# ConvertHeosPlaylistToM3u
Convert HEOS playlist to m3u format

Based on the JSON type playlist (Playlist) obtained with the HEOS command
Generate m3u type playlist with file path information added  
  
[First argument]: JSON type playlist file path acquired by HEOS command  
 Dim PlayListFilePath As String = "D:\tmp\Playlist 1.txt"

[Second argument]: ROOT folder path of media file  
 Dim targetFolder As String = "F:\DATA\"

[Third argument]: Cache file path to save MP3Tag information of media file  
 Dim MP3TagInfoFilePath As String = "F:\tmp\MP3FileTags.txt"  
  
 If the cache file does not exist, from the ROOT folder path of the media file  
 Created based on the information of all media files  
  
  
■ Overview  
1. MP3Tag information of media files included in the specified folder and below
   Collect the path information to each media files in the cache file.
   If the cache file exists, this process is skipped.

2. Generate m3u format playlist from HEOS playlist.
   Compare with the playlist and MP3 Tag information (cache file),
   If it matches, add it to the playlist in m3u format.

   * The output folder of m3u file is directly under the ROOT folder path specified by the second argument.
     Create a folder called #Playlists#.
   * The file name of the m3u file is the playlist file extension specified in the first argument replaced with m3u.

Please refer to the following blog for detailed explanation.  
https://ameblo.jp/nabezou3/entry-12616420663.html  
https://ameblo.jp/nabezou3/entry-12617340479.html  

■ Build environment
 VisualStudio VB.NET Windows Console Project
* Add Reference (COM): Microsoft Shell Controls And Automation
* Addition of reference (assembly extension): Json.NET is a popular high-performance JSON framework for .NET

■ Reference site
* DOBON.NET  
  https://dobon.net/vb/dotnet/file/getabsolutepath.html
* Search or get the files under the folder to the lowest level  
  http://jeanne.wankuma.com/tips/vb.net/directory/getfilesmostdeep.html

■ Correction history
* 2020.08.11 Command line argument launch version  
* 2020.08.11 Change to display end message instead of Console.ReadLine() at the end  
* 2020.08.12 Clarified reference site  
