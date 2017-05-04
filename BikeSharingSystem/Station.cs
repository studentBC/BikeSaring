using System;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BikeSharingSystem
{
    public class Station:IComparable
    {
        private double capacity;
        private int stationID;
        private double initial;//該站初始腳踏車數
        private double Lock ;
        private double rate; //每分鐘減少的比率
        private double declinebymin = 0; //每分鐘減少幾個
        private double surplusbymin = 0; //每分鐘增加幾個
        private double ratePerCapacity= 0;//單位capacity下common goods的增減比率
        public int StationID
        {
            get { return stationID; }
        }
        public double Capacity
        {
            get { return capacity; }
        }
        public double Surplusbymin
        {
            get { return surplusbymin; }
        }
        public double Declinebymin
        {
            get { return declinebymin; }
        }
        public double Locker
        {
            set
            {
                if (this.Lock <= this.capacity && this.Lock >=0)
                {
                    Lock = value;
                }
                else
                {
                    throw new ArgumentException("your lock exceed capacity");
                }
            }
            get { return Lock; }
        }
        public double currentGoods
        {
            set
            {
                if (this.initial <= this.capacity && this.initial>=0)
                {
                    initial = value;
                }
                else
                {
                    throw new ArgumentException("your goods exceed capacity");
                }
            }
            get { return initial; }
        }
        public double Rate
        {
            get { return rate; }
        }
        public double RatePerCapacity
        {
            get { return ratePerCapacity; }
        }
       
        /// <summary>
        /// station constructor 
        /// </summary>
        /// <param name="c">capacity</param>
        /// <param name="i">initial commodity</param>
        /// <param name="r">decline or surplus ratio</param>
        public Station(int c,int i,int id,double r)
        {
            rate = r;
            capacity = c;
            initial = i;
            stationID = id;
            Lock = capacity - initial;
            if(r < 0) { declinebymin = rate ; }
            else if(r >0) { surplusbymin = rate ; }
            ratePerCapacity = Math.Abs(this.rate) / this.capacity;
        }
        /// <summary>
        /// 開始缺少貨物或者多出的時點
        /// </summary>
        /// <param name="starttime">到達該站的時間點</param>
        /// <param name="PnD">到那站後pick up and delivery 的數量</param>
        /// <returns></returns>
        public double LackSurplusTimePoint(double starttime,int PnD)
        {
            if(rate < 0)
            {   //卡車到達該站時該站的腳踏車就沒ㄌ
                if (Lock/declinebymin*-1 < starttime)
                {
                    return initial/declinebymin*-1;
                }
                else
                {
                    initial = Convert.ToInt32( initial + starttime *rate);
                    initial += PnD;
                    return  initial/declinebymin+starttime;
                }
            }
            else
            {
                //卡車到達該站時該站的腳踏車就滿ㄌ
                if (Lock/surplusbymin < starttime)
                {
                    return surplusbymin/declinebymin;
                }
                else
                {
                    initial = Convert.ToInt32(initial + starttime * rate);
                    initial += PnD;
                    return initial/surplusbymin + starttime;
                }
            }
        }
        //每單位容量下的增減比率 值越大越先服務
        public int CompareTo(Station that)
        {
            if (this.ratePerCapacity == that.ratePerCapacity)
            {
                return 0;
            }
            else if (this.ratePerCapacity > that.ratePerCapacity)
            {
                return 1;
            }
            else { return -1; }
        }

        public int CompareTo(object obj)
        {
      
            if (obj == null) return 1;

            Station that = obj as Station;
            if (that != null)
                return this.RatePerCapacity.CompareTo(that.RatePerCapacity);
            else
                throw new ArgumentException("Object is not a station");
        
        }
    }
}
