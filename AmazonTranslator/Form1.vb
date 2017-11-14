Imports System.IO
Imports System.Text
Imports Microsoft.VisualBasic.FileIO

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' Test Mode disables the refreshTimer (loads all orders) and changes the folders.
            Dim test As Boolean = False
            Dim manual As Boolean = False


            Dim srcDirectory As String = "C:\Users\Administrator\amtu2\DocumentTransport\production\reports\"
            Dim destDirectory As String = "C:\eNVenta-ERP\BMECat\Amazon\"
            Dim orderName As String = "ORDER*.txt"
            Dim filename As String = "AmazonOrders.csv"
            Dim refreshTimer As Integer = 360

            If test Then
                destDirectory = "C:\Users\s.ewert\Desktop\"
                srcDirectory = "C:\Users\s.ewert\amtu2\DocumentTransport\production\reports\"
                refreshTimer = 6000000
            End If
            If manual Then
                srcDirectory = "V:\BMECat\Amazon\Test"
                destDirectory = "V:\BMECat\Amazon\"
                refreshTimer = 6000000
            End If

            Dim text_file_path As String = ""
            Dim text_file_path_new As String = ""
            Dim file_dest As String = ""
            Dim file_old As String = ""
            Dim file_data As String

            Dim file_text As String = ""

            Dim dirs As String() = Directory.GetFiles(srcDirectory, orderName)
            Dim filenumber As Integer
            filenumber = dirs.Length

            file_dest = destDirectory + filename
            ' Generate CSV file
            Dim data As System.IO.StreamWriter
            data = My.Computer.FileSystem.OpenTextFileWriter(file_dest, False)

            If filenumber > 0 Then
                For i As Integer = 0 To filenumber - 1
                    text_file_path = dirs(i)

                    ' Get creation date of file
                    Dim dt As DateTime = File.GetLastWriteTime(text_file_path)
                    Dim now As DateTime = DateTime.Now

                    ' It shall only read files from the last 10 minutes
                    Dim timedif As Integer = (now - dt).TotalSeconds

                    If timedif < refreshTimer Then
                        ' Read stuff from txt file
                        Dim tfp As New TextFieldParser(text_file_path, Encoding.GetEncoding(1252))  ' Source files are in ANSI-encoding
                        tfp.Delimiters = New String() {vbTab}
                        tfp.TextFieldType = FieldType.Delimited

                        Dim old_order As String = ""
                        tfp.ReadLine() 'Skips header
                        While tfp.EndOfData = False
                            file_data = tfp.ReadLine()
                            file_data = Replace(file_data, ";", ",")
                            file_data = Replace(file_data, vbTab, ";")
                            Dim file_data_array As String() = file_data.Split(New Char() {";"c})

                            file_data = ""
                            For value As Integer = 0 To file_data_array.Length - 1
                                ' Changes Parameters to shipping parameters
                                Select Case value
                                    Case 11
                                        ' Fix for prices
                                        ' Current price divided by number of articles
                                        Dim new_price As Double = Convert.ToDouble(file_data_array(value)) / 100 / Convert.ToDouble(file_data_array(9))
                                        Dim new_price_string As String = Replace(new_price.ToString, ",", ".")   ' We need a . for decimals
                                        file_data = file_data + new_price_string + ";"
                                    Case 17
                                    ' Do Stuff for 17 in step 18
                                    Case 18
                                        ' User is private, not business
                                        If file_data_array(value).Length = 0 Then
                                            file_data = file_data + file_data_array(value - 1) + ";;"
                                        Else
                                            file_data = file_data + file_data_array(value) + ";" + file_data_array(value - 1) + ";"
                                        End If
                                    Case 25
                                    ' Do Stuff for 17 in step 18
                                    Case 26
                                        ' User is private, not business
                                        If file_data_array(value).Length = 0 Then
                                            file_data = file_data + file_data_array(value - 1) + ";;"
                                        Else
                                            file_data = file_data + file_data_array(value) + ";" + file_data_array(value - 1) + ";"
                                        End If
                                    Case Else
                                        file_data = file_data + file_data_array(value) + ";"
                                End Select
                            Next

                            data.WriteLine(file_data)
                            ' Shipping shouldn't be calculated two times
                            If String.Compare(old_order, file_data_array(0)) Then
                                Dim shipping As String = ""
                                For value As Integer = 0 To file_data_array.Length - 1
                                    ' Changes Parameters to shipping parameters
                                    Select Case value
                                        Case 7
                                            shipping = shipping + "VERSAND-1955_LAGER" + ";"    ' Shipping sku
                                        Case 8
                                            shipping = shipping + "Versand Amazon" + ";"        ' Shipping name
                                        Case 9
                                            shipping = shipping + "1" + ";"                     ' Numbers of shippings
                                        Case 11
                                            shipping = shipping + "4.9" + ";"                   ' Shipping cost
                                        Case 17
                                    ' Do Stuff for 17 in step 18
                                        Case 18
                                            ' User is private, not business
                                            If file_data_array(value).Length = 0 Then
                                                shipping = shipping + file_data_array(value - 1) + ";;"
                                            Else
                                                shipping = shipping + file_data_array(value) + ";" + file_data_array(value - 1) + ";"
                                            End If
                                        Case 25
                                    ' Do Stuff for 17 in step 18
                                        Case 26
                                            ' User is private, not business
                                            If file_data_array(value).Length = 0 Then
                                                shipping = shipping + file_data_array(value - 1) + ";;"
                                            Else
                                                shipping = shipping + file_data_array(value) + ";" + file_data_array(value - 1) + ";"
                                            End If
                                        Case Else
                                            shipping = shipping + file_data_array(value) + ";"
                                    End Select
                                Next
                                ' Write shipping to file
                                data.WriteLine(shipping)
                            End If

                            old_order = file_data_array(0)
                        End While
                    End If
                Next
            Else
                data.WriteLine()
            End If

            data.Close()

            ' Log all exceptions
        Catch ex As Exception
            IO.File.AppendAllText("log.log", String.Format("{0}{1}", Environment.NewLine, DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss") + " " + ex.ToString()))
        End Try

        Application.Exit()
    End Sub
End Class
