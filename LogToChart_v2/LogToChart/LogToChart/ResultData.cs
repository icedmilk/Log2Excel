public class ResultData
{
    public int hour;
    public int minute;
    public double utilAvg;
    public double utilSum;
    public double utilCount;
    public ResultData(int h, int m)
    {
        hour = h;
        minute = m;
        utilCount = 0;
        utilSum = 0;
    }

    public void InsertData(double data)
    {
        utilSum += data;
        utilCount++;
    }

    public void Calc()
    {
        utilAvg = utilSum / utilCount;
    }
}
