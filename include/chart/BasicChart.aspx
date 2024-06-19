<%@ Page language="VB" CodeFile="BasicChart.aspx.vb" Inherits="Chart" CodePage="65001" %>
<html>
    <head>
        <title>WebForm1</title>
        <meta content="Microsoft Visual Studio 7.0" name="GENERATOR"/>
        <meta content="VB" name="CODE_LANGUAGE"/>
        <meta content="JavaScript" name="vs_defaultClientScript"/>
        <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema"/>
        <link media="all" href="chart.css" type="text/css" rel="stylesheet" />
    </head>
    <body style="margin:0px;">
    
    <%
        'Server.UrlDecode(Request.QueryString("gd"))
        'Request.ContentType = "Shift_JIS"
        'Response.Write(Server.UrlDecode(Request.QueryString("gd")))
        'Response.End()
    %>
                   <asp:chart id="Chart1" runat="server" Height="290px" Width="600px" 
                       ImageLocation="/TempImages/ChartPic_#SEQ(300,3)" Palette="BrightPastel" 
                       imagetype="Png" BorderlineDashStyle="Solid" BackSecondaryColor="White" 
                       BackGradientStyle="Center" BorderWidth="2" backcolor="#D3DFF0" 
                       BorderColor="26, 59, 105" Visible="True">
                        <legends>
                            <asp:Legend IsTextAutoFit="False" Name="Default" BackColor="Transparent" Font="Trebuchet MS, 7pt, style=Bold" DockedToChartArea="ChartArea1" Docking="Top" IsDockedInsideChartArea="true" Enabled="true"></asp:Legend>
                        </legends>
                        <borderskin skinstyle="None"></borderskin>
                        <series>
                            <asp:Series Name="Default" IsVisibleInLegend="False">
                                <points>
                                    <asp:DataPoint AxisLabel="‚SŒŽ" />
                                    <asp:DataPoint AxisLabel="‚TŒŽ" />
                                    <asp:DataPoint AxisLabel="‚UŒŽ" />
                                    <asp:DataPoint AxisLabel="‚VŒŽ" />
                                    <asp:DataPoint AxisLabel="‚WŒŽ" />
                                    <asp:DataPoint AxisLabel="‚XŒŽ" />
                                    <asp:DataPoint AxisLabel="‚P‚OŒŽ" />
                                    <asp:DataPoint AxisLabel="‚P‚PŒŽ" />
                                    <asp:DataPoint AxisLabel="‚P‚QŒŽ" />
                                    <asp:DataPoint AxisLabel="‚PŒŽ" />
                                    <asp:DataPoint AxisLabel="‚QŒŽ" />
                                    <asp:DataPoint AxisLabel="‚RŒŽ" />
                                </points>
                            </asp:Series>
                        </series>
                        <chartareas>
                            <asp:ChartArea Name="ChartArea1" BorderColor="64, 64, 64, 64" BorderDashStyle="Solid" BackSecondaryColor="White" BackColor="64, 165, 191, 228" ShadowColor="Transparent" BackGradientStyle="TopBottom" Visible="True" AlignmentStyle="All" BackImageAlignment="TopLeft" AlignmentOrientation="Vertical">
                                <area3dstyle Rotation="10" perspective="10" Inclination="15" IsRightAngleAxes="False" wallwidth="0" IsClustered="False" Enable3D="False"></area3dstyle>
                                <axisy linecolor="64, 64, 64, 64" IntervalType="number" Interval="Auto" ArrowStyle="Triangle" IntervalAutoMode="VariableCount" TitleFont="Trebuchet MS, 8px, style=Bold" TextOrientation="Stacked" IsInterlaced="False" IsMarksNextToAxis="True">
                                    <labelstyle font="Trebuchet MS, 8.25pt, style=Bold" Format="{0:C}" />
                                    <majorgrid linecolor="64, 64, 64, 64" />
                                </axisy>
                                <axisx linecolor="64, 64, 64, 64" LabelAutoFitMaxFontSize="8" Interval="1"  IsMarginVisible="false">
                                    <labelstyle font="Trebuchet MS, 8pt, style=Bold" IsStaggered="false" />
                                    <majorgrid linecolor="64, 64, 64, 64" />
                                </axisx>
                            </asp:ChartArea>
                        </chartareas>
                    </asp:chart>
    </body>
</html>

