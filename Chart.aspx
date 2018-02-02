<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Chart.aspx.cs" Inherits="Chart" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
    <script type="text/javascript">
        var dt = [];
        function testFunction(data) {
            Redraw(data);
        }

        function Redraw(details) {
            // Load the Visualization API and the corechart package.
            google.charts.load('current', {
                'packages': ['corechart']
            });

            // Set a callback to run when the Google Visualization API is loaded.
            google.charts.setOnLoadCallback(drawChart);

            // Callback that creates and populates a data table,
            // instantiates the pie chart, passes in the data and
            // draws it.
            function drawChart() {

                // Create the data table.
                var data = new google.visualization.DataTable();
                data.addColumn('string', 'Dept');
                data.addColumn('number', 'Sales Amount');
                data.addColumn('number', 'Sales Per');
                data.addRows(details);

                var total = google.visualization.data.group(data, [{
                    type: 'boolean',
                    column: 0,
                    modifier: function () { return true; }
                }], [{
                    type: 'number',
                    column: 1,
                    aggregation: google.visualization.data.sum
                }], [{
                    type: 'number',
                    column: 2,
                    aggregation: google.visualization.data.sum
                }]);

                data.addRow(['Total: ' + total.getValue(0, 1), 0, 2]);

                //data.addRows([
                //  ['Department-1', 1000, 7],
                //  ['Department-2', 2000, 13],
                //  ['Department-3', 3000, 20],
                //  ['Department-4', 4000, 27],
                //  ['Department-5', 5000, 33]
                //]);

                // Set chart options
                var options = {
                    'title': 'Department Wise Sales Figure',
                    'width': 600,
                    'height': 400,
                    'sliceVisibilityThreshold': 0
                };

                // Instantiate and draw our chart, passing in some options.
                var chart = new google.visualization.PieChart(document.getElementById('chart_div'));
                chart.draw(data, options);
            }
        }

    </script>
</head>
<body>
    <form id="form1" runat="server">
        <asp:Panel ID="MyPanel" runat="server" HorizontalAlign="Center">
            <asp:Button ID="btnDownload" runat="server" Text="Download Excel" OnClick="DownloadFile_Click" />
            <br />
            <br />
            <br />
            <br />
            <asp:FileUpload ID="FileUpload1" runat="server" />
            <br />
            <br />
            <asp:Button ID="btnUpload" runat="server" Text="Upload" OnClick="btnUpload_Click" />
            <asp:Button ID="btnClear" runat="server" Text="Reset" OnClick="ClearUpload_Click" />
            
            <br />
            <br />

            <div id="chart_div" style="margin-left:30%"></div>
        </asp:Panel>
    </form>
</body>
</html>
