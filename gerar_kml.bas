Sub Gerarkml()

Dim Current As Worksheet
Dim lInha1, iNicio, PlName, Ficheiro As String
Dim TxtIcon As String
Dim iTotaL, aRquivo As Integer
Dim SheetMap As String
Dim Longe As String
Dim Late As String
Dim Color1, Color2, Color3, Color4 As String
Dim lPlanilha As String
Dim CorDaColuna As String
Dim iCont, Cabeca   As Long
Dim IconeCor As Long
Dim ColIconTxt As Long 
Dim ColLog As Long
Dim ColLat As Long 
Dim ColLink As Long
Dim ColLinkDesc As Long 
Dim iColunA, UltimaCol As Long
Dim ColunaLink As Range 
Dim ColunaLinkDesc As Range 
Dim IconColunaCor As Range
Dim ColunaLati As Range
Dim ColunaLonge As Range
Dim ColunaIconTxt As Range

Ficheiro = PastaSistema(16) & "\" & ActiveWorkbook.Name & ".kml"
aRquivo = FreeFile
Open Ficheiro For Output As aRquivo ' Abre o ficheiro para escrita

'Imprime codificacao do xml kml
Print #aRquivo, "<?xml version=""1.0"" encoding=""UTF-8""?><kml xmlns=""http://www.opengis.net/kml/2.2""      xmlns:gx=""http://www.google.com/kml/ext/2.2""><Document><name><![CDATA["
Print #aRquivo, ActiveWorkbook.Name; ""
    
'Imprime descrição, data do arquivo kml com logo Sinal business que será visualizada ao clicar no kml
Print #aRquivo, "]]></name><Snippet maxLines='0'></Snippet><open>1</open><Style><IconStyle><Icon></Icon></IconStyle><BalloonStyle><text><![CDATA[$[description]]]></text></BalloonStyle></Style>" _
& "<description><![CDATA[<img src=""https://static.wixstatic.com/media/b03e29_d86601ab1b134725981557e9447961b6~mv2.png/v1/fill/w_250,h_88/business%20300x100.png"">" _
& "<style type='text/css'>*{font-family:Verdana,Arial,Helvetica,Sans-Serif;}</style><table style=""width: 270px;""><tr><td><br/> </td></tr><tr><td style=""vertical-align: top;"">Arquivo:</td>" _
& "<td style=""width: 100%;"">" & ActiveWorkbook.Name & "</td></tr><tr><td>Data:</td><td>"
Print #aRquivo, VBA.Strings.Format(Now, "dd/mm/yyyy hh:mm")
Print #aRquivo, "UTC <br/></td></tr><tr><td colspan=""2"" style=""vertical-align: center;""><br/>" _
                    & "Uso Restrito - Direitos Reservado. </trd</tr></table>]]></description>"
     
'Imprime mapa de Estilos a ser usada nos placemark correspondete as cores Vermelho, Amarelo, Verde
'Sheet1Map1 AMARELO
Print #aRquivo, "<StyleMap id=""Sheet1Map1""><Pair><key>normal</key><styleUrl>#NormalSheet1Map1</styleUrl></Pair>" _
& "<Pair><key>highlight</key><styleUrl>#HighlightSheet1Map1</styleUrl></Pair></StyleMap>" _
& "<Style id=""NormalSheet1Map1""><IconStyle><scale>1.3</scale><Icon><href>http://www.earthpoint.us/Default.ashx?IconText=&amp;Color=White&amp;File=White.jpg</href></Icon>" _
& "<color>FF00FFFF</color><!-- AMARELO --><hotSpot x=""0.5"" y=""2"" xunits=""fraction"" yunits=""pixels""/></IconStyle>" _
& "<LabelStyle><color>00000000</color><scale>1</scale></LabelStyle>" _
& "<BalloonStyle><text><![CDATA[$[description]]]></text></BalloonStyle></Style>" _
& "<Style id=""HighlightSheet1Map1""><IconStyle><scale>1.6</scale><Icon><href>http://www.earthpoint.us/Default.ashx?IconText=&amp;Color=White&amp;File=White.jpg</href></Icon>" _
& "<color>FF00FFFF</color><!-- AMARELO --><hotSpot x=""0.5"" y=""2"" xunits=""fraction"" yunits=""pixels""/></IconStyle>" _
& "<LabelStyle><color>00000000</color><scale>1</scale></LabelStyle>" _
& "<BalloonStyle><text><![CDATA[$[description]]]></text></BalloonStyle></Style>"
'Sheet1Map2 AZUL CLARO
Print #aRquivo, "<StyleMap id=""Sheet1Map2""><Pair><key>normal</key><styleUrl>#NormalSheet1Map2</styleUrl></Pair>" _
& "<Pair><key>highlight</key><styleUrl>#HighlightSheet1Map2</styleUrl></Pair></StyleMap>" _
& "<Style id=""NormalSheet1Map2""><IconStyle><scale>1.3</scale><Icon><href>http://www.earthpoint.us/Default.ashx?IconText=&amp;Color=White&amp;File=White.jpg</href></Icon>" _
& "<color>FFFFFF00</color><!-- AZUL CLARO   --><hotSpot x=""0.5"" y=""2"" xunits=""fraction"" yunits=""pixels""/></IconStyle>" _
& "<LabelStyle><color>00000000</color><scale>1</scale></LabelStyle><LineStyle><color>FFFF00FF</color><width>2</width></LineStyle>" _
& "<BalloonStyle><text><![CDATA[$[description]]]></text></BalloonStyle></Style><Style id=""HighlightSheet1Map2""><IconStyle><scale>1.6</scale>" _
& "<Icon><href>http://www.earthpoint.us/Default.ashx?IconText=&amp;Color=White&amp;File=White.jpg</href></Icon>" _
& "<color>FFFFFF00</color><!-- AZUL CLARO   --><hotSpot x=""0.5"" y=""2"" xunits=""fraction"" yunits=""pixels""/></IconStyle>" _
& "<LabelStyle><color>00000000</color><scale>1</scale></LabelStyle><LineStyle><color>FFFF00FF</color>" _
& "<width>3</width></LineStyle><BalloonStyle><text><![CDATA[$[description]]]></text></BalloonStyle></Style>"
'Sheet1Map3 VERDE
Print #aRquivo, "<StyleMap id=""Sheet1Map3""><Pair><key>normal</key><styleUrl>#NormalSheet1Map3</styleUrl></Pair>" _
& "<Pair><key>highlight</key><styleUrl>#HighlightSheet1Map3</styleUrl></Pair></StyleMap>" _
& "<Style id=""NormalSheet1Map3""><IconStyle><scale>1.3</scale><Icon><href>http://www.earthpoint.us/Default.ashx?IconText=&amp;Color=White&amp;File=White.jpg</href></Icon>" _
& "<color>FF00FF00</color><!-- VERDE   --><hotSpot x=""0.5"" y=""2"" xunits=""fraction"" yunits=""pixels""/></IconStyle>" _
& "<LabelStyle><color>00000000</color><scale>1</scale></LabelStyle><LineStyle><color>FFFF00FF</color><width>2</width></LineStyle>" _
& "<BalloonStyle><text><![CDATA[$[description]]]></text></BalloonStyle></Style><Style id=""HighlightSheet1Map3""><IconStyle><scale>1.6</scale>" _
& "<Icon><href>http://www.earthpoint.us/Default.ashx?IconText=&amp;Color=White&amp;File=White.jpg</href></Icon>" _
& "<color>FF00FF00</color><!-- VERDE   --><hotSpot x=""0.5"" y=""2"" xunits=""fraction"" yunits=""pixels""/></IconStyle>" _
& "<LabelStyle><color>00000000</color><scale>1</scale></LabelStyle><LineStyle><color>FFFF00FF</color>" _
& "<width>3</width></LineStyle><BalloonStyle><text><![CDATA[$[description]]]></text></BalloonStyle></Style>"
'Sheet1Map4 VERMELHO
Print #aRquivo, "<StyleMap id=""Sheet1Map4""><Pair><key>normal</key><styleUrl>#NormalSheet1Map4</styleUrl></Pair>" _
& "<Pair><key>highlight</key><styleUrl>#HighlightSheet1Map4</styleUrl></Pair></StyleMap><Style id=""NormalSheet1Map4""><IconStyle><scale>1.3</scale><Icon>" _
& "<href>http://www.earthpoint.us/Default.ashx?IconText=&amp;Color=White&amp;File=White.jpg</href></Icon><color>FF0000FF</color><hotSpot x=""0.5"" y=""2"" xunits=""fraction"" yunits=""pixels""/>" _
& "</IconStyle><LabelStyle><color>00000000</color><scale>1</scale></LabelStyle><LineStyle><color>FFFF00FF</color><width>2</width></LineStyle><BalloonStyle><text><![CDATA[$[description]]]></text>" _
& "</BalloonStyle></Style><Style id=""HighlightSheet1Map4""><IconStyle><scale>1.6</scale><Icon><href>http://www.earthpoint.us/Default.ashx?IconText=&amp;Color=White&amp;File=White.jpg</href></Icon>" _
& "<color>FF0000FF</color><!--  VERMELHO  --><hotSpot x=""0.5"" y=""2"" xunits=""fraction"" yunits=""pixels""/></IconStyle><LabelStyle><color>00000000</color><scale>1</scale></LabelStyle><LineStyle>" _
& "<color>FFFF00FF</color><width>3</width></LineStyle><BalloonStyle><text><![CDATA[$[description]]]></text></BalloonStyle></Style>"
'Sheet1Map5 LARANJA
Print #aRquivo, "<StyleMap id=""Sheet1Map5""><Pair><key>normal</key>" _
& "<styleUrl>#NormalSheet1Map5</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#HighlightSheet1Map5</styleUrl></Pair>" _
& "</StyleMap><Style id=""NormalSheet1Map5""><IconStyle><scale>1.3</scale><Icon><href>http://www.earthpoint.us/Default.ashx?IconText=&amp;Color=White&amp;File=White.jpg</href></Icon>" _
& "<color>FF1478F0</color><!--LARANJA--><hotSpot x=""0.5"" y=""2"" xunits=""fraction"" yunits=""pixels""/></IconStyle><LabelStyle><color>00000000</color><scale>1</scale></LabelStyle>" _
& "<BalloonStyle><text><![CDATA[$[description]]]></text></BalloonStyle></Style><Style id=""HighlightSheet1Map5""><IconStyle>" _
& "<scale>1.6</scale><Icon><href>http://www.earthpoint.us/Default.ashx?IconText=&amp;Color=White&amp;File=White.jpg</href></Icon>" _
& "<color>640A78F0</color><!--LARANJA--> <hotSpot x=""0.5"" y=""2"" xunits=""fraction"" yunits=""pixels""/></IconStyle><LabelStyle><color>00000000</color><scale>1</scale></LabelStyle><LineStyle><color>FFFF00FF</color>" _
& "<width>3</width></LineStyle><BalloonStyle><text><![CDATA[$[description]]]></text></BalloonStyle></Style>"
'Sheet1Map6 BRANCA
Print #aRquivo, "<StyleMap id=""Sheet1Map6""><Pair><key>normal</key><styleUrl>#NormalSheet1Map6</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#HighlightSheet1Map6</styleUrl></Pair>" _
& "</StyleMap><Style id=""NormalSheet1Map6""><IconStyle><scale>1.3</scale><Icon><href>http://www.earthpoint.us/Default.ashx?IconText=&amp;Color=White&amp;File=White.jpg</href>" _
& "</Icon><color>White</color><hotSpot x=""0.5"" y=""2"" xunits=""fraction"" yunits=""pixels""/></IconStyle><LabelStyle><color>00000000</color><scale>1</scale></LabelStyle>" _
& "<LineStyle><color>FFFF00FF</color><width>2</width></LineStyle><BalloonStyle><text><![CDATA[$[description]]]></text></BalloonStyle></Style><Style id=""HighlightSheet1Map6""><IconStyle>" _
& "<scale>1.6</scale><Icon><href>http://www.earthpoint.us/Default.ashx?IconText=&amp;Color=White&amp;File=White.jpg</href></Icon><color>White</color>" _
& "<hotSpot x=""0.5"" y=""2"" xunits=""fraction"" yunits=""pixels""/></IconStyle><LabelStyle><color>00000000</color><scale>1</scale></LabelStyle><LineStyle><color>FFFF00FF</color>" _
& "<width>3</width></LineStyle><BalloonStyle><text><![CDATA[$[description]]]></text></BalloonStyle></Style>"

'Sheet1Map7 AZUL ESCURO
Print #aRquivo, "<StyleMap id=""Sheet1Map7""><Pair><key>normal</key>" _
& "<styleUrl>#NormalSheet1Map7</styleUrl></Pair><Pair><key>highlight</key><styleUrl>#HighlightSheet1Map7</styleUrl></Pair>" _
& "</StyleMap><Style id=""NormalSheet1Map7""><IconStyle><scale>1.3</scale><Icon><href>http://www.earthpoint.us/Default.ashx?IconText=&amp;Color=White&amp;File=White.jpg</href></Icon>" _
& "<color>F4F00014</color><!--AZUL--><hotSpot x=""0.5"" y=""2"" xunits=""fraction"" yunits=""pixels""/></IconStyle><LabelStyle><color>00000000</color><scale>1</scale></LabelStyle>" _
& "<BalloonStyle><text><![CDATA[$[description]]]></text></BalloonStyle></Style><Style id=""HighlightSheet1Map7""><IconStyle>" _
& "<scale>1.6</scale><Icon><href>http://www.earthpoint.us/Default.ashx?IconText=&amp;Color=White&amp;File=White.jpg</href></Icon>" _
& "<color>F4F00014</color><!--AZUL--> <hotSpot x=""0.5"" y=""2"" xunits=""fraction"" yunits=""pixels""/></IconStyle><LabelStyle><color>00000000</color><scale>1</scale></LabelStyle><LineStyle><color>FFFF00FF</color>" _
& "<width>3</width></LineStyle><BalloonStyle><text><![CDATA[$[description]]]></text></BalloonStyle></Style>"

'Legenda topo Sinal Business
'https://static.wixstatic.com/media/b03e29_6fa8712df0d44841b8da8836b87b50c4~mv2.png  - LOGO AZUL
Print #aRquivo, "<ScreenOverlay><name>Sinal Business</name><Icon><href>https://static.wixstatic.com/media/b03e29_abca4d48868b45dd921c4ca4ed0b8f48~mv2.png/v1/fill/w_1707,h_702/Sinal%20Business%202.png</href></Icon>" _
& "<overlayXY x=""0"" y=""1"" xunits=""fraction"" yunits=""fraction""/><screenXY x=""0"" y=""1"" xunits=""fraction"" yunits=""fraction""/><rotationXY x=""00"" y=""00"" xunits=""fraction"" yunits=""fraction""/>" _
& "<size x=""256"" y=""105"" xunits=""pixels"" yunits=""pixels""/></ScreenOverlay>"



'VARIAVEIS DAS CORES DA LINHA
    Color1 = "WHITE"
    Color2 = "#ddffdd" 'verde claro
    Color3 = "#f7d56d"  'AMARELO SINAL
    Color4 = "'#c5d9f1'" ''AZUL CLARO


Cabeca = 2  'linha do cabeçario da planilha
lPlanilha = ActiveSheet.Name 'pega o nome da planilha ativa

'INICIO DA IMPRESSAO DA FOLHA DE PLANILHA

'LOOP para percorrer as Worksheets
For Each Current In Worksheets

Print #aRquivo, "<Folder><name><![CDATA[" & Worksheets(Current.Name).Name & "]]></name><visibility>1</visibility><open>0</open>"

                    Worksheets(Current.Name).Activate 


iTotaL = Cells(Rows.Count, 1).End(xlUp).Row  ' conta as linhas usadas de baixo para cima
                        
    iCont = Cabeca 'Cabeca se refere a linha de Cabeçario que deve ser descontada no loop
        
              'LOOP DAS LINHAS
            While iCont < iTotaL
                iCont = iCont + 1
                            
'DEFINI A COLUNA DO ICONCOR , que não poderá ter seu nome alterado
Set IconColunaCor = Range("A2:AZ2").Find("IconCor") 'Busca a coluna de Iconcor, se nao encontrar gera erro
                    IconeCor = IconColunaCor.Column

Set ColunaLati = Range("A2:AZ2").Find("Latitude") 'Busca a coluna de Latitude,  se nao encontrar gera erro
                    ColLat = ColunaLati.Column
                            
Set ColunaLonge = Range("A2:AZ2").Find("Longitude") 'Busca a coluna de Longitude,  se nao encontrar gera erro
                    ColLog = ColunaLonge.Column
                            
Set ColunaLink = Range("A2:AZ2").Find("Link:") 'Busca a coluna de Link,  se nao encontrar gera erro
                    ColLink = ColunaLink.Column

Set ColunaLinkDesc = Range("A2:AZ2").Find("Link Descritivo") 'Busca a coluna de Link,  se nao encontrar gera erro
                    ColLinkDesc = ColunaLinkDesc.Column
                                                       
Set ColunaIconTxt = Range("A2:AZ2").Find("IconText") 'Busca a coluna de Longitude,  se nao encontrar gera erro
                    ColIconTxt = ColunaIconTxt.Column
                                                      
             TxtIcon = Cells(iCont, ColIconTxt)
               Longe = Left(Cells(iCont, ColLog).Text, 10) ' Coluna Longitude, remove o grau
                Late = Left(Cells(iCont, ColLat).Text, 10) ' Coluna Latitude, remove o grau
              
              'DETERMINA A COR DO ICONE
                    If Cells(iCont, IconeCor) = "amarelo" Then
                        SheetMap = "#Sheet1Map1"
                        ElseIf Cells(iCont, IconeCor) = "verde" Then
                        SheetMap = "#Sheet1Map3"
                        ElseIf Cells(iCont, IconeCor) = "vermelho" Then
                        SheetMap = "#Sheet1Map4"
                        ElseIf Cells(iCont, IconeCor) = "laranja" Then
                        SheetMap = "#Sheet1Map5"
                        ElseIf Cells(iCont, IconeCor) = "branca" Then
                        SheetMap = "#Sheet1Map6"
                        ElseIf Cells(iCont, IconeCor) = "azul" Then
                        SheetMap = "#Sheet1Map7"
                        Else
                        SheetMap = "#Sheet1Map2"
                    End If
                    
            ' APARENCIA DO ICONE - COLUNAS A, Z(latitude),  AB(longitude), COR ICON E ICON TEXT
            
            Print #aRquivo, "<Placemark><name><![CDATA[" & Cells(iCont, 1) & "]]></name><Snippet maxLines='0'></Snippet><styleUrl>" & SheetMap & "</styleUrl>" _
                            & "<Style><IconStyle><Icon><href>http://www.earthpoint.us/Default.ashx?IconText=" & TxtIcon & "&amp;Color=White&amp;File=White57654302.jpg" _
                            & "</href></Icon></IconStyle></Style><LookAt><longitude>" & Longe & "</longitude><latitude>" & Late & "</latitude><range>1000</range></LookAt>" _
                            & "<Point><coordinates>" & Longe & "," & Late & "</coordinates></Point>"
                        
            
            'COLUNA A NOME DO ICONE "DESCRITIVO  - sem  white-space: nowrap;
            Print #aRquivo, "<description><![CDATA[<table border='1' cellspacing='0' cellpadding='1'><tr bgcolor=" & Color4 & "><td colspan='2' style='text-align:center; padding-left:" _
                            & "10px; padding-right: 10px;'><b>" & Cells(iCont, 1).Text & "</b></td></tr>"
            
            UltimaCol = Cells(2, Cells.Columns.Count).End(xlToLeft).Column - 6
            
               CorDaColuna = Color1
            'LOOP DAS COLUNAS
          
             For iColunA = 2 To UltimaCol
            
            
            'COLUNA MODELO -  white-space: nowrap;
            Print #aRquivo, "<tr bgcolor=" & CorDaColuna & "><td style='vertical-align: top; padding-left: 10px; white-space: nowrap;'><b>" & Cells(Cabeca, iColunA) & "</b></td>" _
                            & "<td style='vertical-align: top; padding-left: 6px; padding-right: 10px;'>" & Cells(iCont, iColunA).Text & "</td></tr>"
                
            'Altera cor da coluna
            If CorDaColuna = Color1 Then
            CorDaColuna = Color2
            Else
            CorDaColuna = Color1
            End If
            'FECHA LOOP DAS COLUNAS
        Next
        
            'COLUNA Link e Link Descritivo
            Print #aRquivo, "<tr bgcolor=" & Color3 & "><td style='vertical-align: top; padding-left: 10px; white-space: nowrap;'><b>" & Cells(Cabeca, ColLink) & "</b></td>" _
                            & "<td style='vertical-align: top; padding-left: 6px; padding-right: 10px; width: 600px;'><a target='_blank' href=" & Cells(iCont, ColLink) & "'><b>" & Cells(iCont, ColLinkDesc) & "</b></a></td></tr>"
     
             'COLUNA DE FECHAMENTO DO ITEM
            Print #aRquivo, "<tr bgcolor=" & Color3 & "><td colspan='2' style='vertical-align: top; padding-left:30px; padding-right: 10px; white-space: nowrap;'><i>Direitos Reservados - Sinal Business</i></td></tr>"
            
            'Fecha O ITEM DA LINHA
            Print #aRquivo, "</table>]]></description></Placemark>"
            
    
        Wend 'VAI PARA PROXIMA LINHA DA SHEET

            'Fecha a folha da Planilha
            Print #aRquivo, "</Folder>"

   
Next ' VAI PARA A PROXIMA SHEET DA PLANILHA

            'fecha o kml
           Print #aRquivo, "</Document></kml>"
            
            'fechando arquivo
            Close aRquivo
            
Worksheets(lPlanilha).Activate 'retorna a planilha ativa
            
            MsgBox "O Arquivo: " & ActiveWorkbook.Name & ".kml Foi exportado para sua Área de Tabalho"

End Sub