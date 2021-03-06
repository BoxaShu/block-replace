﻿'' bx_blockreplace.vb
'' © Шульжицкий Владимир, 2014 (boxa.shu@gmail.com)
'' Назначение: Плагин для AutoCAD, предназначен для добавление атрибутов в блоки .
'' Заказчик: Александр Чумак <ca@vodeco.ru> <zvezdo4et@list.ru> ООО «ВОДЭКО»  www.vodeco.ru
'' Команда: bx_blockreplace

Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Windows

Public Class acad__boxashu
    Const CrLf As String = ControlChars.CrLf 'Environment.NewLine'ControlChars.CrLf
    Private keyWord As String


    <CommandMethod("bx_blockreplace")> _
    Public Sub bx_blockreplace()

        Me.keyWord = "ABout"

        '' Получениеn текущего документа и базы данных
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim acEd As Editor = acDoc.Editor

        Dim pm As ProgressMeter = New ProgressMeter()
        'тут лежит список с именами блоков которые нужно поменять
        Dim listBlocksToReplace As New List(Of String)

        'Список файлов и пути к ним
        Dim listBlock As New Dictionary(Of String, String)

        '' Старт транзакции
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            'Выбираем только атрибуты
            Dim acTypValAr(0) As TypedValue
            acTypValAr.SetValue(New TypedValue(DxfCode.Start, "INSERT"), 0)
            '' Назначение критериев фильтра объекту SelectionFilter
            Dim acSelFtr As SelectionFilter = New SelectionFilter(acTypValAr)
            'http://www.theswamp.org/index.php?PHPSESSID=avlhqv9o52tm7kiauq62c122j2&topic=31864.0;nowap
            Dim opt As New PromptSelectionOptions()
            opt.SetKeywords("[ABout]", "ABout")
            opt.MessageForAdding = CrLf & "Выберите блоки для замены или " & opt.Keywords.GetDisplayString(True)
            AddHandler opt.KeywordInput, AddressOf onKeywordInput

            '' Запрос выбора объектов в области чертежа
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection(opt, acSelFtr)

            '' Если статус запроса равен OK, объекты выбраны
            If acSSPrompt.Status <> PromptStatus.OK Then
                Exit Sub
            End If


            Dim acSSet As SelectionSet = acSSPrompt.Value
            '' Перебор объектов в наборе
            For Each acSSObj As SelectedObject In acSSet
                '' Проверка, нужно убедится в правильности полученного объекта
                If Not IsDBNull(acSSObj) Then
                    '' Открытие объекта для записи
                    Dim acEnt As Entity = CType(acTrans.GetObject(acSSObj.ObjectId, _
                                                            OpenMode.ForRead), Entity)
                    'Dim acPline As Polyline = CType(acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead), Polyline)
                    If Not IsDBNull(acEnt) Then
                        If TypeOf acEnt Is BlockReference Then
                            Dim acBlock As BlockReference = CType(acEnt, BlockReference)
                            If listBlocksToReplace.Contains(acBlock.Name) = False Then
                                listBlocksToReplace.Add(acBlock.Name)
                            End If
                        End If
                    End If
                End If
            Next
            ' Сохранение нового объекта в базе данных
            acTrans.Commit()
            ' Очистка транзакции
        End Using
        acEd.WriteMessage(CrLf & "Выбрано блоков: {0}", listBlocksToReplace.Count)

        If listBlocksToReplace.Count > 0 Then
            Dim openPathDialog As New Windows.Forms.FolderBrowserDialog()
            openPathDialog.RootFolder = Environment.SpecialFolder.DesktopDirectory
            ' openPathDialog.RootFolder = Environment.SpecialFolder.NetworkShortcuts
            openPathDialog.ShowDialog()
            Dim PATH As String = openPathDialog.SelectedPath
            If PATH <> "" Then
                acEd.WriteMessage(CrLf & "Выбран каталог {0}", PATH)
            Else
                acEd.WriteMessage(CrLf & "Ошибка при выборе каталога! Программа завершена.")
                Exit Sub
            End If


            Dim files As String()
            Try
                files = System.IO.Directory.GetFiles(PATH, "*.dwg", IO.SearchOption.AllDirectories)
            Catch ex As System.ArgumentException
                acEd.WriteMessage(CrLf & "Ошибка при поиске файлов! Программа завершена.")
                acEd.WriteMessage(CrLf & ex.Message)
                Exit Sub
            End Try


            If files.Count > 0 Then
                pm.SetLimit(files.Count)
                pm.Start("Формирую список файловю")
                For Each f As String In files
                    Dim fil As System.IO.FileInfo = New System.IO.FileInfo(f)
                    Dim fileName As String = fil.Name.Substring(0, fil.Name.Length - fil.Extension.Length)
                    'acEd.WriteMessage(CrLf & "Доступны файлы {0}, путь до файла {1}", fileName, f)
                    listBlock.Add(f, fileName)
                    pm.MeterProgress()
                Next
                pm.Stop()
            Else
                acEd.WriteMessage(CrLf & "В выбранном каталоге нет dwg файлов! Программа завершена.")
                Exit Sub
            End If

            acEd.WriteMessage(CrLf & "В каталоге найдено файлов с блоками: {0}", listBlock.Count)

        Else
            acEd.WriteMessage(CrLf & "Блоки для замены не выбраны! Программа завершена.")
            Exit Sub
        End If



        '' тут читаем dwg файл с блоками и копируем в описание блока атрибуты из файла
        pm.SetLimit(listBlocksToReplace.Count)
        pm.Start("Изменение блоков")
        Dim correctBlock As Integer = 0

        For Each i As String In listBlocksToReplace
            Dim blockID As ObjectId
            Dim attDic As New Dictionary(Of String, String)
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acBlockTable As BlockTable = CType(acTrans.GetObject(acCurDb.BlockTableId, _
                                                               OpenMode.ForRead), BlockTable)
                Dim acBlockTableRecord As BlockTableRecord = CType(acTrans.GetObject(acBlockTable.Item(i), _
                                                               OpenMode.ForWrite), BlockTableRecord)

                If acBlockTableRecord.HasAttributeDefinitions = True Then
                    For Each id As ObjectId In acBlockTableRecord
                        Dim obj As DBObject = acTrans.GetObject(id, OpenMode.ForRead)
                        If TypeOf obj Is AttributeDefinition Then
                            Dim ent1 As AttributeDefinition = CType(obj, AttributeDefinition)

                            If attDic.ContainsKey(ent1.Tag) <> True Then
                                attDic.Add(ent1.Tag, ent1.TextString)
                            End If
                        End If
                    Next
                End If

                blockID = acBlockTableRecord.ObjectId
                ' Сохранение нового объекта в базе данных
                acTrans.Commit()
                ' Очистка транзакции
            End Using


            Try
                Dim izm As Boolean = False
                If listBlock.ContainsValue(i) = True Then

                    Dim SourcePath As String = (From q In listBlock Where q.Value = i Select q.Key).First.ToString
                    Using dbSource As Autodesk.AutoCAD.DatabaseServices.Database = New Database(False, True)
                        dbSource.ReadDwgFile(SourcePath, System.IO.FileShare.Read, True, "")

                        'get the model space object ids for both dbs
                        Dim sourceMsId As ObjectId = SymbolUtilityServices.GetBlockModelSpaceId(dbSource)

                        'тут надо вставить ID описания блока
                        'Dim destDbMsId As ObjectId = SymbolUtilityServices.GetBlockModelSpaceId(acCurDb)
                        Dim destDbMsId As ObjectId = blockID


                        'now create an array of object ids to hold the source objects to copy
                        Dim sourceIds As ObjectIdCollection = New ObjectIdCollection()

                        Using tm As Transaction = dbSource.TransactionManager.StartTransaction()
                            Dim bt As BlockTable = CType(tm.GetObject(dbSource.BlockTableId, OpenMode.ForRead, False), BlockTable)
                            Dim btr As BlockTableRecord = CType(tm.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead, False), BlockTableRecord)
                            For Each id As ObjectId In btr
                                Try
                                    'http://adn-cis.org/forum/index.php?topic=806.0
                                    'Ускоряем перебор
                                    'If id.ObjectClass.Name = GetType(AttributeDefinition).Name Then
                                    'End If
                                    'Вариант с наследованием
                                    'Dim dimenClass As RXClass = RXObject.GetClass(GetType(AttributeDefinition))
                                    'If id.ObjectClass.IsDerivedFrom(dimenClass) Then
                                    'End If

                                    Dim ent1 As Entity = CType(tm.GetObject(id, OpenMode.ForRead), Entity)
                                    If Not IsDBNull(ent1) Then
                                        If TypeOf ent1 Is AttributeDefinition Then

                                            'sourceIds.Add(id)

                                            Dim ent0 As AttributeDefinition = CType(ent1, AttributeDefinition)
                                            If attDic.ContainsKey(ent0.Tag) = True Then
                                                acEd.WriteMessage(CrLf & "В блоке {0}, выявлен конфликт атрибутов {1}." & _
                                                                 " Добавление данного атрибута пропущено.", i, ent0.Tag)
                                            Else
                                                'Dim acAtt As AttributeDefinition = CType(ent1, AttributeDefinition)
                                                'acAttrCallection.Add(acAtt)
                                                sourceIds.Add(id)
                                            End If

                                        End If
                                    End If

                                    'тут нужно добавить проверку блока, который меняем, 
                                    'на наличие в нем аттрибутов и их значения

                                Catch ex As Autodesk.AutoCAD.Runtime.Exception
                                    acEd.WriteMessage(CrLf & "Ошибка при чтении блока из файла {0}! Программа завершена.", SourcePath)
                                    acEd.WriteMessage(CrLf & ex.Message)
                                    Exit Sub
                                End Try
                            Next
                            'next prepare to deepclone the recorded ids to the destdb
                            If sourceIds.Count > 0 Then
                                Dim mapping As IdMapping = New IdMapping()
                                'now clone the objects into the destdb
                                dbSource.WblockCloneObjects(sourceIds, destDbMsId, mapping, DuplicateRecordCloning.Replace, False)
                                izm = True
                            End If

                            tm.Commit()
                        End Using
                    End Using

                End If

                If izm Then
                    correctBlock = correctBlock + 1
                End If


            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                acEd.WriteMessage(CrLf & "Ошибка при изменении блока {0}! Программа завершена.", i)
                acEd.WriteMessage(CrLf & ex.Message)
                Exit Sub
            End Try
            pm.MeterProgress()
        Next
        pm.Stop()
        acEd.WriteMessage(CrLf & "____________________")
        acEd.WriteMessage(CrLf & "Изменено блоков: {0}", correctBlock)


 

    End Sub

    Private Sub onKeywordInput(sender As Object, e As SelectionTextInputEventArgs)

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acEd As Editor = acDoc.Editor

        'Me.keyWord = e.Input
        'acEd.WriteMessage(ControlChars.CrLf & "Вы ввели: " & e.Input)


        acEd.WriteMessage(CrLf & "Программа: bx_blockreplace.")
        acEd.WriteMessage(CrLf & "Предназначена для сравнения блоков в чертеже, " _
                               & "с блоками в указанной папке и импортирования атрибутов из блоков " _
                               & "в папке, в блоки чертежа, при условии совпадения имен блоков. " _
                               & "Если атрибуты имеются и в блоках чертежа и в блоках в папке то, " _
                               & "при совпадении имен атрибутов, данный атрибут не будет " _
                               & "импортирован и об этом будет выведено в консоль автокада " _
                               & "соответствующее предупреждение.")
        acEd.WriteMessage(CrLf & "При обнаружении неточности или ошибки просьба" _
                               & " сообщить автору, по адресу: armspec@gmail.com")
        acEd.WriteMessage(CrLf & "Для коммерческого использования." _
                               & " Лицензия на программу принадлежит: ООО «ВОДЭКО»  www.vodeco.ru")
        acEd.WriteMessage(CrLf & "Количество рабочих мест: неограничено.")
        acEd.WriteMessage(CrLf & "Стандартное соглашение" _
                                  & " тут: http://experement.spb.ru/wiki/doku.php?id=programming:lic_pay")
        acEd.WriteMessage(CrLf & " _ ")
        acEd.WriteMessage(CrLf & " _ ")
        acEd.WriteMessage(CrLf & " _ ")
        acEd.WriteMessage(CrLf & " _ ")
        acEd.WriteMessage(CrLf & " _ ")

    End Sub
End Class
