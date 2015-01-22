'*****************************************************************************************
'                           LICENSE INFORMATION
'*****************************************************************************************
'   SD Split & Merge version 1.0
'   A simple software for split & Merge any file easily

'   Copyright (C) 2014  
'   Sajjad Tanha
'   Email : sajjadtanha20@gmail.com
'   Created: 20Dec14

'   This program is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program.  If not, see <http:'www.gnu.org/licenses/>.
'*****************************************************************************************

Imports System.IO
Public Class SplitMerge

    Private _FileName As String
    Private _ChunkSize As Integer
    Private _Error As String
    Private _OutPutPath As String
    Private _DeleteFileAfterSplit As Boolean
    Private _DeleteFilesAfterMerge As Boolean
    Public Event FileSplitCompleted(ByVal ErrorMessage As String)
    Public Event FileMergeCompleted(ByVal ErrorMessage As String)
    Public Event UpdateProgress(ByVal Progress As Double)

    Public Property DeleteFilesAfterMerge() As Boolean
        Get
            Return _DeleteFilesAfterMerge
        End Get
        Set(ByVal Value As Boolean)
            _DeleteFilesAfterMerge = Value
        End Set
    End Property

    Public Property DeleteFileAfterSplit() As Boolean
        Get
            Return _DeleteFileAfterSplit
        End Get
        Set(ByVal Value As Boolean)
            _DeleteFileAfterSplit = Value
        End Set
    End Property

    Public Property OutputPath() As String
        Get
            Return _OutPutPath
        End Get
        Set(ByVal Value As String)
            If Value.Length > 0 AndAlso Value.Substring(Value.Length - 1, 1) = "\" Then Value = Value.Substring(0, Value.Length - 1)
            _OutPutPath = Value
        End Set
    End Property

    Public ReadOnly Property getError() As String
        Get
            Return _Error
        End Get
    End Property

    Public Sub ClearErrors()
        _Error = ""
    End Sub

    Public Property FileName() As String
        Get
            Return _FileName
        End Get
        Set(ByVal Value As String)
            If Value <> "" Then
                If Value.Substring(Value.Length - 4, 4) = ".001" Then
                    _FileName = Value.Substring(0, Value.Length - 4)
                Else
                    _FileName = Value
                End If
            End If
        End Set
    End Property

    Public Property ChunkSize() As Integer
        Get
            Return _ChunkSize
        End Get
        Set(ByVal Value As Integer)
            _ChunkSize = Value
        End Set
    End Property

    Public Sub SplitFile()
       
        Dim _FileSize As Long
        Dim _Index As Long
        Dim _OutputFile As String
        Dim _BaseName As String
        Dim _StartPosition As Long
        Dim _Buffer As Byte() = New Byte() {}
        Dim _InputFileStram As System.IO.FileStream
        Dim _OutputFileStram As System.IO.FileStream
        Dim _BinaryWriter As BinaryWriter
        Dim _BinaryReader As BinaryReader
        Dim _Fragments As Long
        Dim _RemainingBytes As Long
        Dim _SplitFileName As String
        Dim _Progress As Double

        Try
            _Error = ""
            _SplitFileName = Me.FileName
            _InputFileStram = New FileStream(FileName, FileMode.Open)
            _BinaryReader = New BinaryReader(_InputFileStram)
            _FileSize = FileSize(_SplitFileName)
            _BaseName = FileBaseName(_SplitFileName)
            If Not File.Exists(_SplitFileName) Then _Error = "file " & FileName & " Not Exists."
            If _FileSize <= ChunkSize Then _Error = _SplitFileName & " Chunk size (" & ChunkSize & ") is bigger than original file size(" & _FileSize & ")"
            If _Error = "" Then
                _Fragments = _FileSize / ChunkSize
                _Progress = 100 / _Fragments
                _RemainingBytes = _FileSize - (_Fragments * ChunkSize)
                If Me.OutputPath = "" Then Me.OutputPath = Directory.GetParent(_SplitFileName).ToString
                If Not Directory.Exists(Me.OutputPath) Then Directory.CreateDirectory(Me.OutputPath)
                _BinaryReader.BaseStream.Seek(0, SeekOrigin.Begin)
                For _Index = 1 To _Fragments
                    _OutputFile = Me.OutputPath & "\" & _BaseName & "." & Format(_Index, "000")
                    ReDim _Buffer(ChunkSize - 1)
                    _BinaryReader.Read(_Buffer, 0, ChunkSize)
                    _StartPosition = _BinaryReader.BaseStream.Seek(0, SeekOrigin.Current)
                    If File.Exists(_OutputFile) Then File.Delete(_OutputFile)
                    _OutputFileStram = New System.IO.FileStream(_OutputFile, FileMode.Create)
                    _BinaryWriter = New BinaryWriter(_OutputFileStram)
                    _BinaryWriter.Write(_Buffer)
                    _OutputFileStram.Flush()
                    _BinaryWriter.Close()
                    _OutputFileStram.Close()
                    RaiseEvent UpdateProgress(_Progress)
                Next
                If _RemainingBytes > 0 Then
                    _OutputFile = Me.OutputPath & "\" & _BaseName & "." & Format(_Index, "000")
                    ReDim _Buffer(_RemainingBytes - 1)
                    _BinaryReader.Read(_Buffer, 0, _RemainingBytes)
                    If File.Exists(_OutputFile) Then File.Delete(_OutputFile)
                    _OutputFileStram = New System.IO.FileStream(_OutputFile, FileMode.Create)
                    _BinaryWriter = New BinaryWriter(_OutputFileStram)
                    _BinaryWriter.Write(_Buffer)
                    _OutputFileStram.Flush()
                    _BinaryWriter.Close()
                    _OutputFileStram.Close()
                End If
                _InputFileStram.Close()
                _BinaryReader.Close()
                If Me.DeleteFileAfterSplit Then File.Delete(FileName)
                Me._Error = ""
            End If
        Catch ex As Exception
            Me._Error = ex.ToString
        Finally
            _BinaryWriter = Nothing
            _OutputFileStram = Nothing
            _BinaryReader = Nothing
            _InputFileStram = Nothing
            RaiseEvent UpdateProgress(100)
            RaiseEvent FileSplitCompleted(Me._Error)
        End Try
    End Sub

    Public Sub MergeFile()
        Dim _InputFileStram As System.IO.FileStream
        Dim _OutputFileStram As System.IO.FileStream
        Dim _BinaryWriter As BinaryWriter
        Dim _BinaryReader As BinaryReader
        Dim _MergeFiles() As String
        Dim _index As Short
        Dim _buffer() As Byte = New Byte() {}
        Dim _FileSize As Long
        Dim _MergedFile As String
        Dim _MergingFileName As String
        Dim _Progress As Double
        Try
            _Error = ""
            _MergingFileName = Me.FileName
            If Me.FileName = "" Then _Error = "Filename should not be empty."
            If _Error = "" Then
                _MergeFiles = getMergeFiles(_MergingFileName)
                _Progress = 100 / (_MergeFiles.Length - 1)
                If Not IsNothing(_MergeFiles) Then
                    If Me.OutputPath = "" Then Me.OutputPath = Directory.GetParent(_MergingFileName).ToString
                    If Not Directory.Exists(Me.OutputPath) Then Directory.CreateDirectory(Me.OutputPath)
                    _MergedFile = Me.OutputPath & "\" & FileBaseName(_MergingFileName)
                    If File.Exists(_MergedFile) Then File.Delete(_MergedFile)
                    _OutputFileStram = New FileStream(_MergedFile, FileMode.CreateNew)
                    _BinaryWriter = New BinaryWriter(_OutputFileStram)
                    For _index = 0 To _MergeFiles.Length - 1
                        _FileSize = FileSize(_MergeFiles(_index))
                        _InputFileStram = New FileStream(_MergeFiles(_index), FileMode.Open)
                        _BinaryReader = New BinaryReader(_InputFileStram)
                        _BinaryReader.BaseStream.Seek(0, SeekOrigin.Begin)
                        ReDim _buffer(_FileSize - 1)
                        _BinaryReader.Read(_buffer, 0, _buffer.Length)
                        _BinaryWriter.Write(_buffer)
                        _OutputFileStram.Flush()
                        _BinaryReader.Close()
                        _InputFileStram.Close()
                        RaiseEvent UpdateProgress(_Progress)
                    Next
                    _OutputFileStram.Close()
                    _BinaryWriter.Close()
                    If Me.DeleteFilesAfterMerge Then
                        For _index = 0 To _MergeFiles.Length - 1
                            File.Delete(_MergeFiles(_index))
                        Next
                    End If
                Else
                    _Error = ".001 file not found in this path " & vbNewLine & Me.OutputPath
                End If
            End If
        Catch ex As Exception
            Me._Error = ex.ToString
        Finally
            _BinaryWriter = Nothing
            _OutputFileStram = Nothing
            _BinaryReader = Nothing
            _InputFileStram = Nothing
            RaiseEvent UpdateProgress(100)
            RaiseEvent FileMergeCompleted(Me._Error)
        End Try
    End Sub

    Public Function FileSize(ByVal FileName As String) As Long
        Dim _fileInfo As System.IO.FileInfo
        _fileInfo = New FileInfo(FileName)
        Return _fileInfo.Length
        _fileInfo = Nothing
    End Function

    Private Function FileBaseName(ByVal FileName As String) As String
        Dim _fileInfo As System.IO.FileInfo
        _fileInfo = New FileInfo(FileName)
        Return _fileInfo.Name
        _fileInfo = Nothing
    End Function

    Private Function getMergeFiles(ByVal MergeFileName As String) As String()
        Dim _tempFiles As String() = System.IO.Directory.GetFiles(System.IO.Directory.GetParent(MergeFileName).ToString)
        Dim _MergeFiles() As String = Nothing
        Dim _Index As Short
        Do
            _Index += 1
            If Array.BinarySearch(_tempFiles, MergeFileName & "." & Format(_Index, "000")) >= 0 Then
                ReDim Preserve _MergeFiles(_Index - 1)
                _MergeFiles(_Index - 1) = MergeFileName & "." & Format(_Index, "000")
            Else
                Exit Do
            End If
        Loop
        Return _MergeFiles
    End Function

End Class
'''

