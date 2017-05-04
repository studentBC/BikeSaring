using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BikeSharingSystem
{
    public class Truck
    {
        private int capacity;
        private int currentGoods;
        private int currentSpace;
        public int CurrentGoods
        {
            set
            {
                if (currentGoods <= capacity && currentGoods >= 0)
                {
                    currentGoods = value;
                }
                else
                {
                    throw new Exception("bug!!!");
                }
            }
            get { return currentGoods; }
        }
        public int CurrentSpace
        {
            set
            {
                if (currentSpace <= capacity && currentGoods >= 0)
                {
                    currentSpace = value;
                }
                else
                {
                    throw new Exception("bug!!!");
                }
            }
            get { return currentSpace; }
        }

        //private double startTime;
        public Truck(int c,int initialamount)
        {
            capacity = c;
            currentGoods = initialamount;
            currentSpace = capacity - currentGoods;
        }
        int temp = -1;
        public int SupplyTruck(Station s)
        {
            int half = this.capacity / 2;
            if (this.currentGoods == 0)
            {
                if(s.currentGoods >= half)
                {
                    s.currentGoods -= half;
                    this.currentSpace = this.currentGoods = (int)half;
                    s.Locker = Convert.ToInt32(s.Capacity - s.currentGoods);
                    return half;
                }
                else
                {
                    temp = this.currentGoods = (int)s.currentGoods;
                    this.currentSpace = this.capacity - this.currentGoods;
                    s.currentGoods = 0;
                    s.Locker = (int)s.Capacity;
                    return temp;
                }
            }
            else if(this.currentSpace == 0)
            {
                if (s.Locker >= half)
                {
                    s.Locker -= half;
                    this.currentSpace  = this.currentGoods = (int)half;
                    s.currentGoods = Convert.ToInt32(s.Capacity - s.Locker);
                    return half;
                }
                else
                {
                    temp = (int)s.Locker;
                    this.currentSpace += (int)s.Locker;
                    this.currentGoods = capacity - this.currentSpace;
                    s.currentGoods = (int)s.Capacity;
                    s.Locker = 0;
                    return temp;
                }
            }
            return 0;
        }
        double ans = -1;
        public int PickupDelivery(int amount)
        {
            if (amount < 0) //might need to pickup
            {
                if (currentSpace > 0)//貨車還能放的情況下
                {
                    if (-1 * amount <= currentSpace) //放得下
                    {
                        this.currentGoods -=  amount;
                        this.currentSpace = this.capacity - this.currentGoods;
                        return amount;
                    }
                    else //放不下
                    {
                        ans = this.currentSpace * -1;
                        this.currentSpace = 0;
                        this.currentGoods = this.capacity;
                        return (int)ans;
                    }
                }
            }
            else if (amount > 0) //need to delivery
            {
                if (currentGoods > 0)//貨車還能送貨的情況下
                {
                    if (currentGoods >= amount)//可完全補足
                    {
                        this.currentGoods -= amount;
                        this.currentSpace = this.capacity - this.currentGoods;
                        return amount;
                    }
                    else //無法完全補足
                    {
                        ans = this.currentGoods;
                        this.currentGoods = 0;
                        this.currentSpace = this.capacity;
                        return (int)ans;
                    }
                }
            }
            return 0;
        }

    }
}
