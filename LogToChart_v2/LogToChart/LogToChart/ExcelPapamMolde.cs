using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel; 

public class ExcelPapamMolde
{
    /// <summary>
    /// 绑定X轴开始行
    /// </summary>
    public int XstartRow { get; set; }


    /// <summary>
    /// 绑定X轴结束行
    /// </summary>
    public int XendRow { get; set; }


    /// <summary>
    /// 绑定x轴的数据源的列
    /// </summary>
    public int XColumn { get; set; }


    /// <summary>
    /// 绑定报表图形的开始行
    /// </summary>
    public int ValuestartRow { get; set; }


    /// <summary>
    /// 绑定报表图形的结束行
    /// </summary>
    public int ValueendRow { get; set; }


    /// <summary>
    /// 绑定报表图形的列
    /// </summary>
    public int ValueColumn { get; set; }


    /// <summary>
    /// 报表名称
    /// </summary>
    public string ChartName { get; set; }


    /// <summary>
    /// x轴名
    /// </summary>
    public string XAxisName { get; set; }


    /// <summary>
    ///  y轴名
    /// </summary>
    public string YAxisName { get; set; }


    /// <summary>
    /// 图形高度
    /// </summary>
    public int PicHeight { get; set; }


    /// <summary>
    /// 图形宽度
    /// </summary>
    public int PicWidth { get; set; }


    /// <summary>
    /// 页名
    /// </summary>
    public string PageName { get; set; }


    /// <summary>
    /// 图距excel上端距离
    /// </summary>
    public int Top { get; set; }


    /// <summary>
    /// 图距excel左端距离
    /// </summary>
    public int Left { get; set; }
}

/// <summary>
/// 操作EXCEl的chart 封转的类
/// </summary>
public class ExcelControlChart
{
    Workbook m_workBook;
    Worksheet m_workSheet;


    //轴标题偏移量
    const int m_titleOffset = 20;
    //CHART的范围偏移
    const int CHARTAREAEXCURSION = 10;
    //设置绘图区宽度
    const int PLOTAREAWIDTH = 400;
    //X轴标题字体大小
    const int XAXISTITLEFONTSIZE = 10;
    //X轴坐标轴字体大小
    const int XAXISTICKLABELSFONTSIZE = 10;
    //Y轴标题字体大小
    const int YAXISTITLEFONTSIZE = 10;
    //Y轴坐标轴字体大小
    const int YAXISTICKLABELSFONTSIZE = 10;
    //ChartGroup的宽度
    const int CHARTGROUPWIDTH = 100;


    public ExcelControlChart(Workbook workBook, Worksheet workSheet)
    {
        this.m_workBook = workBook;
        this.m_workSheet = workSheet;
    }
    /// <summary>
    /// 创建一个有底色的标题栏
    /// </summary>
    /// <param name="papam"></param>
    /// <returns></returns>
    public void CreateTitle(int row, int columnbegin, int columnend, string title)
    {
        //画一个有底色的标题栏
        m_workSheet.get_Range(m_workSheet.Cells[row, columnbegin], m_workSheet.Cells[row, columnend]).Interior.ColorIndex = ColorIndex.Shamrock;
        //指定单元格下边框线粗细,和色彩
        m_workSheet.get_Range(m_workSheet.Cells[row, 1], m_workSheet.Cells[row, 1]).Borders.get_Item(XlBordersIndex.xlEdgeBottom).Weight = XlBorderWeight.xlMedium;
        m_workSheet.get_Range(m_workSheet.Cells[row, 1], m_workSheet.Cells[row, 1]).Borders.get_Item(XlBordersIndex.xlEdgeBottom).ColorIndex = ColorIndex.Black;
        this.SetCells(row, 1, title);
    }
    /// <summary>
    /// 创建一个线形报表
    /// </summary>
    /// <param name="papam"></param>
    /// <returns></returns>
    public bool CreateLine3D(ExcelPapamMolde papam)
    {
        //添加一个页面
        m_workBook.Charts.Add(Type.Missing, Type.Missing, 1, Type.Missing);
        //CreateTitle(int row, int columnbegin, int columnend, string title);
        //定义图类型
        m_workBook.ActiveChart.ChartType = XlChartType.xlLine;
        //图数据源(根据列的索引得到当前绑定列的名称)
        m_workBook.ActiveChart.SetSourceData(m_workSheet.get_Range(GetExcelColumnName(papam.ValueColumn) + papam.ValuestartRow
            , GetExcelColumnName(papam.ValueColumn) + papam.ValueendRow), XlRowCol.xlColumns);
        //图形宽和高 
        m_workBook.ActiveChart.ChartArea.Width = papam.PicWidth;
        m_workBook.ActiveChart.ChartArea.Height = papam.PicHeight;


        //表示图示画在SHEET1的，改成自己的SHEET名就好
        //m_workSheet.Name = papam.PageName;
        m_workBook.ActiveChart.Location(XlChartLocation.xlLocationAsObject, papam.PageName);


        //没有这个标题就出不来
        m_workBook.ActiveChart.HasTitle = true;
        //报表名称
        m_workBook.ActiveChart.ChartTitle.Text = papam.ChartName;


        m_workBook.ActiveChart.ChartArea.Width = papam.PicWidth + CHARTAREAEXCURSION;
        m_workBook.ActiveChart.ChartArea.Height = papam.PicHeight + CHARTAREAEXCURSION;


        //图形距离左上角的距离
        m_workBook.ActiveChart.ChartArea.Top = papam.Top;
        m_workBook.ActiveChart.ChartArea.Left = papam.Left;


        #region - 定义绘图区 -
        //设置绘图区的背景色 
        m_workBook.ActiveChart.PlotArea.Interior.ColorIndex = ColorIndex.LightViridity;
        //设置绘图区边框线条


        m_workBook.ActiveChart.PlotArea.Border.LineStyle = XlLineStyle.xlLineStyleNone;
        //设置绘图区宽度
        m_workBook.ActiveChart.PlotArea.Width = PLOTAREAWIDTH;

        #endregion


        #region - 定义X轴 -
        //轴样式
        Axis xAxis = (Axis)m_workBook.ActiveChart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
        //Axis xAxis = (Axis)m_workBook.ActiveChart.Axes(XlAxisType.xlSeriesAxis);
        //x轴显示的值
        xAxis.CategoryNames = m_workSheet.get_Range(GetExcelColumnName(papam.XColumn) + papam.XstartRow
            , GetExcelColumnName(papam.XColumn) + papam.XendRow);


        xAxis.HasTitle = true;
        xAxis.AxisTitle.AutoScaleFont = false; //不关掉自动缩放的话后面的字体大小无法设置
        xAxis.AxisTitle.Font.Size = XAXISTITLEFONTSIZE; //X轴标题字体大小
        xAxis.AxisTitle.Text = papam.XAxisName;//X轴名
        xAxis.TickLabels.AutoScaleFont = false;
        xAxis.TickLabels.Font.Size = XAXISTICKLABELSFONTSIZE; //X轴坐标轴字体大小
        xAxis.AxisTitle.Left = papam.PicWidth - m_titleOffset;


        #endregion


        #region - 定义Y轴 -


        Axis yAxis = (Axis)m_workBook.ActiveChart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);


        yAxis.HasTitle = true;
        yAxis.AxisTitle.AutoScaleFont = false; //不关掉自动缩放的话后面的字体大小无法设置
        yAxis.AxisTitle.Font.Size = YAXISTITLEFONTSIZE; //Y轴标题字体大小
        yAxis.AxisTitle.Text = papam.YAxisName;//Y轴名
        yAxis.TickLabels.AutoScaleFont = false;
        yAxis.TickLabels.Font.Size = YAXISTICKLABELSFONTSIZE; //Y轴坐标轴字体大小
        yAxis.AxisTitle.Top = m_titleOffset;
        #endregion


        //设置绘图区的数据标志（就是线形顶上出现值）显示出值来
        m_workBook.ActiveChart.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, false, false
            , false, false, false, true, false, false, false);


        ChartGroup grp = (ChartGroup)m_workBook.ActiveChart.ChartGroups(1);
        grp.GapWidth = CHARTGROUPWIDTH;


        m_workBook.ActiveChart.PlotArea.Width = papam.PicWidth - m_titleOffset; //设置绘图区宽度
        m_workBook.ActiveChart.PlotArea.Top = m_titleOffset;
        m_workBook.ActiveChart.PlotArea.Height = papam.PicHeight - m_titleOffset; //设置绘图区高度
        m_workBook.ActiveChart.PlotArea.Left = m_titleOffset;


        ////设置Legend图例的位置和格式
        m_workBook.ActiveChart.HasLegend = false;


        return true;


    }
    /// <summary>
    /// 得到Excel的列名
    /// </summary>
    /// <param name="ColumnIndex">列的索引</param>
    /// <returns>列名</returns>
    private string GetExcelColumnName(int ColumnIndex)
    {
        //列的默认个数
        const int cellWidth = 26;
        string[] cellNames = new string[cellWidth + 1] {"", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"
                , "M", "N", "O", "P", "Q", "R", "S" ,"T","U","V","W","X","Y","Z"};
        int cellItems = 0;
        int cellItem = 0;
        cellItems = ColumnIndex / cellWidth;
        cellItem = ColumnIndex % cellWidth;


        if (cellItem == 0)
        {
            cellItems--;
            cellItem = ColumnIndex;
        }
        if (cellItems < 1)
        {
            cellItems = 0;
        }
        string str = cellNames[cellItems] + cellNames[cellItem];
        return str;
    }
    /// <summary>
    /// 向单元格写入数据，对当前WorkSheet操作
    /// </summary>
    /// <param name="rowIndex">行索引</param>
    /// <param name="columnIndex">列索引</param>
    /// <param name="text">要写入的文本值</param>
    public void SetCells(int rowIndex, int columnIndex, string text)
    {
        m_workSheet.Cells[rowIndex, columnIndex] = text;
    }


    /// <summary>
    /// 向单元格写入数据，对当前WorkSheet操作
    /// </summary>
    /// <param name="rowIndex">行索引</param>
    /// <param name="columnIndex">列索引</param>
    /// <param name="text">要写入的文本值</param>
    public void SetMultiCells(int rowIndex, int columnIndex, List<float> data)
    {
        for (int i = 0; i < data.Count; ++i)
        {
            m_workSheet.Cells[rowIndex + i, columnIndex] = data[i];
        }
    }
}

public enum ColorIndex
{
    Black = 1,
    Shamrock = 33, //天蓝 
    LightViridity = 34,//浅青绿
}
