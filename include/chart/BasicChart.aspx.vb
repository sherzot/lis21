Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Web
Imports System.Web.SessionState
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports System.Web.UI.DataVisualization.Charting
Partial Class Chart
	Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Len(Request("gd")) > 0 Then
            Dim i As Int16
            Dim m As Int16
            Dim gd = Split(HttpUtility.HtmlDecode(Request("gd")), ",")

            Dim LineName As String

            Dim ShiftJISEncoding As Encoding = Encoding.GetEncoding("Shift_JIS")

            Dim UnicodeEncoding As Encoding = Encoding.Unicode

            Dim ShiftJisBytes As Byte()

            Dim bytesData As Byte()

            Dim series As Object

            For i = 0 To UBound(gd) Step 14
                LineName = gd(i + 0)

                ShiftJisBytes = ShiftJISEncoding.GetBytes(LineName)

                bytesData = Encoding.GetEncoding("Shift_JIS").GetBytes(LineName)
                'bytesData = Encoding.UTF8.GetBytes(LineName)

                bytesData = System.Text.Encoding.Convert(ShiftJISEncoding, UnicodeEncoding, ShiftJisBytes)

                'series = New Series(LineName)
                If IsNothing(Chart1.Series.FindByName(System.Text.Encoding.Unicode.GetString(bytesData))) = False Then
                    m = 2
                    Do While IsNothing(Chart1.Series.FindByName(System.Text.Encoding.Unicode.GetString(bytesData) + m.ToString())) = False
                        m = m + 1
                    Loop
                    series = New Series(System.Text.Encoding.Unicode.GetString(bytesData) + m.ToString())
                Else
                    series = New Series(System.Text.Encoding.Unicode.GetString(bytesData))
                End If

                series.ChartType = SeriesChartType.FastLine
                series.BorderWidth = 2
                series.ShadowOffset = 2
                series.LabelBorderWidth = 2
                series.Points.AddY(gd(i + 1)) '4ŒŽ
                series.Points.AddY(gd(i + 2)) '5ŒŽ
                series.Points.AddY(gd(i + 3)) '6ŒŽ
                series.Points.AddY(gd(i + 4)) '7ŒŽ
                series.Points.AddY(gd(i + 5)) '8ŒŽ
                series.Points.AddY(gd(i + 6)) '9ŒŽ
                series.Points.AddY(gd(i + 7)) '10ŒŽ
                series.Points.AddY(gd(i + 8)) '11ŒŽ
                series.Points.AddY(gd(i + 9)) '12ŒŽ
                series.Points.AddY(gd(i + 10)) '1ŒŽ
                series.Points.AddY(gd(i + 11)) '2ŒŽ
                series.Points.AddY(gd(i + 12)) '3ŒŽ
                'series.Color = System.Drawing.Color.Red
                series.Color = ColorTranslator.FromHtml("0x" & gd(i + 13))
                'series.Color = ColorTranslator.FromHtml("0xFF0000")
                Chart1.Series.Add(series)
            Next
        End If

        ' Create new data series and set it's visual attributes
        '		Dim series As New Series("Spline")
        'series.ChartType = SeriesChartType.Spline
        '		series.ChartType = SeriesChartType.FastLine

        '		series.BorderWidth = 3
        '		series.ShadowOffset = 2

        ' Populate new series with data
        '		series.Points.AddY(67)
        '		series.Points.AddY(57)
        '		series.Points.AddY(83)
        '		series.Points.AddY(23)
        '		series.Points.AddY(70)
        '		series.Points.AddY(60)
        '		series.Points.AddY(90)
        '		series.Points.AddY(20)
        ' Add series into the chart's series collection
        '		Chart1.Series.Add(series)

    End Sub 'Page_Load

End Class