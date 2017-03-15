using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;

namespace BikeSharingSystem
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            mypen = new Pen(Color.Blue, 1.5f);
            mypen.CustomEndCap = new System.Drawing.Drawing2D.AdjustableArrowCap(0, 6);
            myfont = new Font("Arial", 10.0f);
        }
        System.Diagnostics.Stopwatch time = new System.Diagnostics.Stopwatch(); //time of each run
        System.Diagnostics.Stopwatch Totaltime = new System.Diagnostics.Stopwatch();  //time of total run
        public static double MovingTime = 1;
        PermutationGA pga = null;
        double[,] distanceMatrix;
        double[] startTOvertex;
        string[] initialGoods;
        int[] PandD;
        int[] BestRouting; //其實是因為懶得再decode一次 所以用額外的記憶體印出答案
        double timehorizon;
        double totalDistance = 0;
        double speed;
        double NotDoAnything = 0;
        double CPUBurstTime = 0.0;
        double PDVelocity; //time cost to pickup and delivery one bike in one minute
        int truckcapacity;
        int totalStation;
        int finishStation;
        int soFarfinishStation; //so far the best solution complete station amount
        double[] longitude;
        double[] latitude;
        double[] BUSA;
        double[] BestUSA;
        Station[] stations;

        //繪圖需要的參數
        double xmin = 0.0, ymax = 0.0, scale = 0.0, scale2 = 0.0, xmiddle = 0.0, ymiddle = 0.0;
        int w = 0, h = 0;
        double W = 0, H = 0;
        Pen mypen;
        Font myfont;
        private void reset(object sender, EventArgs e)
        {
            //pga = new PermutationGA(totalStation*2+2, OptimizationType.Min,ComputeTimeAverageHoldingShortageLevel,stations);
            //timehorizon = Convert.ToDouble(txbtimehorizon.Text.ToString());
            speed = Convert.ToDouble(txbspeed.Text.ToString());
            truckcapacity = Convert.ToInt32(txbcapacity.Text.ToString());
            pga.TruckCapacity = truckcapacity;
            timehorizon = Convert.ToDouble(txbtimehorizon.Text);
            pga.TimeHorizon = timehorizon;
            PDVelocity = 1/Convert.ToDouble(textBoxPnDTime.Text);
            CPUBurstTime = Convert.ToDouble(txbCPUTime.Text);
            if (double.IsInfinity(PDVelocity)) { PDVelocity = double.MaxValue; }
            //pga.TruckCapacity = truckcapacity;
            //pga.TimeHorizon = timehorizon;
            if (checkBoxGreedy.Checked == true)
            {
                pga.Greedyinitialize(distanceMatrix, startTOvertex);
            }
            else
            {
                pga.reset();
            }

            for(int i = 0; i < stations.Length; i++)
            {
                stations[i].currentGoods = Convert.ToInt32(initialGoods[i]);
                stations[i].Locker = (int)stations[i].Capacity - stations[i].currentGoods;
            }
             foreach (Series s in chart.Series)
             {
                 s.Points.Clear();
             }
            richTextBox.Clear();
            min = Double.MaxValue;
            labelBO.Text = "BestObjective  ";
            NotDoAnything = 0;
            Truck uselesstruck = new Truck(truckcapacity, 0);
            for (int i = 0; i < totalStation; i++)
            {
                if (stations[i].Surplusbymin > 0)
                {
                    NotDoAnything += HoldingShortageSurplus(timehorizon, i, 0, ref uselesstruck,truckcapacity);
                }
                else
                {
                    NotDoAnything += HoldingShortageSurplus(timehorizon, i, 0, ref uselesstruck, 0);
                }
            }
            //update data grid view


        }

        private void createPGA(object sender, EventArgs e)
        {
            butreset.Enabled = true;
            butrunend.Enabled = true;
            pga = new PermutationGA(totalStation*2+2, OptimizationType.Min, new GA<int>.ObjectiveFunctionDelegate(ComputeTimeAverageHoldingShortageLevel),stations,distanceMatrix);
            this.ppg.SelectedObject = this.pga;
        }
        
        private void oneIteration(object sender, EventArgs e)
        {
            pga.executeOneIteration();
            chart.Series[0].Points.AddXY((double)pga.IterationCount, pga.IterationAverage);
            chart.Series[1].Points.AddXY((double)pga.IterationCount, pga.IterationBestObjective);
            chart.Series[2].Points.AddXY((double)pga.IterationCount, pga.SoFarTheBestObjective);
         
            double unsatisfiedAmount = Math.Round( pga.SoFarTheBestObjective,2);
            labelBO.Text = "BestObjective  " + Convert.ToString(unsatisfiedAmount);
            labelBO.Text.ToString();
            richTextBox.Text = "繞站順序: depot " ;
            for(int i = 0; i < soFarfinishStation+1; i++)
            {
                //我忘記解碼fuck fuck fuck fuck fuck fuck !! 
                //   richTextBox.Text += pga.SoFarTheBestSolution[i].ToString()+" ";             
                richTextBox.Text += BestRouting[i].ToString() + ",";
            }
            if (soFarfinishStation == totalStation - 1)
            {
                richTextBox.Text += "depot";
            }
            double totalPandD = 0;
            richTextBox.Text +="\n"+"pickup & delivery amount: ";
            for(int i = 0; i <= soFarfinishStation; i++)
            {
                richTextBox.Text += PandD[i].ToString() + ",";
                totalPandD += Math.Abs(PandD[i]);
            }
           
            richTextBox.Text += "\n" + "truck start time: " + pga.SoFarTheBestSolution[totalStation * 2 + 1];
            richTextBox.Text += "\n" + "truck initial common goods: " + pga.SoFarTheBestSolution[totalStation * 2];
            totalDistance = 0;
            // totalDistance += startTOvertex[pga.SoFarTheBestSolution[0]]; //抵達第一站的距離
            //string dist = null;
            totalDistance += startTOvertex[BestRouting[0]]; //抵達第一站的距離
            //dist = totalDistance.ToString()+"->";
            for (int i = 0; i < soFarfinishStation; i++)
            {
                totalDistance += distanceMatrix[BestRouting[i], BestRouting[i + 1]];
                //dist += distanceMatrix[BestRouting[i], BestRouting[i + 1]].ToString() + "->";
            }
            //全部跑完要回去depot
            if (soFarfinishStation == totalStation - 1)
            {
                totalDistance += startTOvertex[BestRouting[soFarfinishStation]];
                //dist += startTOvertex[BestRouting[soFarfinishStation]].ToString() + "->";
            }
            richTextBox.Text +="\n"+ "*******performance*******";
            richTextBox.Text += "\n" + "Total pickup & delivery : " + (totalPandD+ pga.SoFarTheBestSolution[totalStation * 2]);
            richTextBox.Text += "\n" + "total distance :" + totalDistance;
            richTextBox.Text += "\n" + "unsatisfication amount :" + unsatisfiedAmount;
            richTextBox.Text += "\n"+"without action :"+NotDoAnything;
            richTextBox.Text += "\n" + "improvement rate :" + Math.Round(1-unsatisfiedAmount/NotDoAnything,2);
            //richTextBox.AppendText("routing distance "+dist);


            labelMax.Text ="Max :" + NotDoAnything.ToString();
            chart.Series[3].Points.AddXY((double)pga.IterationCount, NotDoAnything);

            //chart.Update();
            //讓圖形跑超過50代次時可跑的較緩慢
            //if ((pga.IterationCount > 50) && (!pga.terminationConditionMet())) { Cursor = System.Windows.Forms.Cursors.WaitCursor; }
            //else { Cursor = System.Windows.Forms.Cursors.Default; }
        }


        private void runToEnd(object sender, EventArgs e)
        {
            if (checkBox_times.Checked == true)
            {
                if (pga.IterationCount == 0)
                {
                    Totaltime.Reset();
                    Totaltime.Start();
                }
                //time.Reset();
                //time.Start();
                while (true)
                {
                    if (Totaltime.Elapsed.TotalSeconds >= CPUBurstTime)
                    {
                        Totaltime.Stop();
                       // time.Stop();
                        break;
                    }
                    else
                    {
                        oneIteration(null, null);
                    }
                }
                labelIteration.Text = "Iterations: " + pga.IterationCount;
            }
            else
            {
                for (int i = 0; i < pga.IterationLimit; i++)
                {
                    oneIteration(null, null);
                }
            }
            ppg.Update();
            if (checkBox.Checked == true)
            {
                for (int c = 0; c <= finishStation; c++)
                {
                    dataGridView.Rows[1].Cells[BestRouting[c]].Value = BestUSA[c];
                    dataGridView.Rows[1].Cells[BestRouting[c]].Style.BackColor = Color.Gray;
                }
                for (int c = finishStation + 1; c < totalStation; c++)
                {
                    dataGridView.Rows[1].Cells[BestRouting[c]].Value = BestUSA[c];
                    dataGridView.Rows[1].Cells[BestRouting[c]].Style.BackColor = Color.Orange;
                }
            }
        }
        double min = Double.MaxValue;
        
        private double ComputeTimeAverageHoldingShortageLevel(int[] chromosomes)
        {

            Truck truck = new Truck(truckcapacity, chromosomes[totalStation * 2]);
            //period為卡車起始出發時間
            double USA = 0; //USA = unsatisfied amount
            double period = chromosomes[totalStation * 2 + 1]+ chromosomes[totalStation * 2]/PDVelocity; //利用distancematrix除以車速即可得到現在的時間
            double totalUSA = 0; 
            totalDistance = 0;
            int[] routingSequence = Decode(chromosomes);
            //the following 5 line is for debug
            //int[] routingSequence = new int[totalStation];
            //for (int a = 0; a < totalStation; a++)
            //{
            //    routingSequence[a] = chromosomes[a];
            //}
            period += (startTOvertex[routingSequence[0]] / speed);//抵達第一站的時間 
       //   totalDistance += startTOvertex[routingSequence[0]]; //抵達第一站的距離
            int i  = 0; int[] temp = new int[totalStation]; finishStation = -1;
            while (period <= timehorizon && i <= totalStation - 1)
            {
                if (stations[routingSequence[i]].Declinebymin < 0 )
                {
                    temp[i] = truck.CurrentGoods;
                    USA = HoldingShortageSurplus(period, routingSequence[i], truck.PickupDelivery(chromosomes[routingSequence[i] + totalStation]), ref truck,temp[i]);
                    temp[i] -= truck.CurrentGoods;
                }
                else if (stations[routingSequence[i]].Surplusbymin > 0 )
                {
                    temp[i] = truck.CurrentSpace;
                    USA = HoldingShortageSurplus(period, routingSequence[i], truck.PickupDelivery(chromosomes[routingSequence[i] + totalStation]), ref truck,temp[i]);
                    temp[i] = truck.CurrentSpace - temp[i];
                }
                else if(stations[routingSequence[i]].Rate == 0 && chromosomes[routingSequence[i] + totalStation] >= 0)
                {
                    temp[i] = truck.CurrentGoods;
                    USA = HoldingShortageSurplus(period, routingSequence[i], truck.PickupDelivery(chromosomes[routingSequence[i] + totalStation]), ref truck, temp[i]);
                    temp[i] -= truck.CurrentGoods;
                }
                else if(stations[routingSequence[i]].Rate == 0 && chromosomes[routingSequence[i] + totalStation] <= 0)
                {
                    temp[i] = truck.CurrentSpace;
                    USA = HoldingShortageSurplus(period, routingSequence[i], truck.PickupDelivery(chromosomes[routingSequence[i] + totalStation]), ref truck, temp[i]);
                    temp[i] = truck.CurrentSpace - temp[i];
                }
                totalUSA += USA;
                BUSA[i] = USA;
                finishStation = i;  ////幹他媽的 bug 在此 幹我找了一個禮拜阿幹!!!
                if (i < totalStation - 1)
                {
                    period += ((distanceMatrix[routingSequence[i], routingSequence[i + 1]] / speed) + Math.Abs(temp[i] / PDVelocity));//到下一站的時間點
                }else { period += Math.Abs(temp[i] / PDVelocity); }
                i++;             
                //update datagrid view for debugging
                if (checkBox.Checked == true)
                {
                    dataGridView.Rows[0].Cells[routingSequence[i]].Value = USA;
                    dataGridView.Rows[0].Cells[routingSequence[i]].Style.BackColor = Color.Gray;
                }
            }



            if (finishStation != totalStation - 1)
            {
                //沒有跑完的車站繼續算成本
                truck.CurrentGoods = 0; truck.CurrentSpace = truckcapacity;
                for (int a = finishStation+1; a < totalStation ; a++)
                {
                    if (stations[routingSequence[a]].Surplusbymin > 0)
                    {
                        USA = HoldingShortageSurplus(timehorizon, routingSequence[a], 0, ref truck, truckcapacity);
                    }
                    else if (stations[routingSequence[a]].Declinebymin < 0)
                    {
                        USA = HoldingShortageSurplus(timehorizon, routingSequence[a], 0, ref truck, 0);
                    }else
                    {
                        USA = 0;
                    }
                    totalUSA += USA;
                    BUSA[a] = USA;
                    if (checkBox.Checked == true)
                    {
                        dataGridView.Rows[0].Cells[routingSequence[a]].Value = USA;
                        dataGridView.Rows[0].Cells[routingSequence[a]].Style.BackColor = Color.Orange;
                    }
                }
            }
            if(min > totalUSA)
            {
                min = totalUSA;
                soFarfinishStation = finishStation;
                temp.CopyTo(PandD,0);
                BUSA.CopyTo(BestUSA,0);
                routingSequence.CopyTo(BestRouting, 0);
            }
            return totalUSA;     
        }
        //其實該list的第0個的值是第1站被繞行的順序
        private int[]Decode(int[] chromosome)
        {
            int[] routingSequence = new int[totalStation];
            int[] temp = new int[totalStation];
            for(int i = 0; i < totalStation; i++)
            {
                temp[i] = chromosome[i];routingSequence[i] = i ; //when you add depot location just amend it
            }
            Array.Sort(temp, routingSequence);
            return routingSequence;
        }
        /// <summary>
        /// 其實還沒考慮過初始站點的bike amount為0時的情況,如果有會fail
        /// </summary>
        /// <param name="period">卡車出發時點</param>
        /// <param name="stationID">station id</param>
        /// <param name="commodity">GA求解後pick up or delivery的數量</param>
        /// <returns></returns>
        private double HoldingShortageSurplus(double starttime,int stationID ,int commodity,ref Truck truck,double TGS)
        {
            double bikeAmount = 0;
            double accumulatedtime = 0;
            double PnDTime = 0;
          //  Truck t = truck;
            if(stations[stationID].Declinebymin < 0) //deliver bicycle to the station commodity must be positive
            {
                double lackTime = 0.0;
                //卡車到達該站時該站的腳踏車就沒ㄌ
                if ( stations[stationID].currentGoods / stations[stationID].Declinebymin * -1 < starttime)
                {
                    lackTime = stations[stationID].currentGoods / stations[stationID].Declinebymin * -1;
                    accumulatedtime = starttime - lackTime;
                    bikeAmount += stations[stationID].Declinebymin * accumulatedtime*-1;
                    //commodity超過locker就要先塞回去不管後續的借還車= =凸操你媽的幹!!
                    if (commodity > stations[stationID].Capacity)
                    {
                        //把多出的或塞回去改變卡車上的數量
                        truck.CurrentGoods += (commodity - (int)stations[stationID].Capacity);
                        truck.CurrentSpace = truckcapacity - truck.CurrentGoods;
                        commodity = (int)stations[stationID].Capacity;
                    }
                    if (bikeAmount < 0) throw new Exception("bug!!!");
                       PnDTime = commodity / PDVelocity;
                    if (starttime + PnDTime > timehorizon) { PnDTime = timehorizon - starttime; }
                    if (PDVelocity + stations[stationID].Declinebymin < 0) //station decline rate > truck supply velocity
                    {  
                        bikeAmount -= (PDVelocity + stations[stationID].Declinebymin) * PnDTime;
                        stations[stationID].currentGoods = 0;
                        stations[stationID].Locker = stations[stationID].Capacity - stations[stationID].currentGoods;
                        if (PnDTime < 0) throw new Exception("bug!!!");
                    }
                    else
                    {
                        if (stations[stationID].Capacity >= (PDVelocity + stations[stationID].Declinebymin) * PnDTime)
                        {
                            stations[stationID].currentGoods = (PDVelocity + stations[stationID].Declinebymin) * PnDTime;
                        }
                        else//如果給的數量超過站的容量就G囉
                        {
                            stations[stationID].currentGoods = stations[stationID].Capacity;
                            truck.CurrentGoods += (commodity - (int)stations[stationID].Capacity);
                            truck.CurrentSpace = truckcapacity - truck.CurrentGoods;
                        }
                        stations[stationID].Locker = stations[stationID].Capacity - stations[stationID].currentGoods;
                      //  PnDTime = (TGS - truck.CurrentGoods) / PDVelocity;
                        if (PnDTime < 0) throw new Exception("bug!!!");
                    }
                        if (stations[stationID].Locker < 0) { throw new Exception("bug is here !!"); }
                        if (commodity != 0)
                        {
                            lackTime = starttime + PnDTime - stations[stationID].currentGoods / stations[stationID].Declinebymin;//又缺為0時
                            if (timehorizon >= lackTime)
                            {
                                bikeAmount += stations[stationID].Declinebymin * (timehorizon - lackTime) * -1;
                            }
                        }
                        else
                        {
                            if (timehorizon >= lackTime)
                            {
                                bikeAmount = (timehorizon - lackTime) * stations[stationID].Declinebymin * -1;
                            }
                        }
                }
                else
                {
                    stations[stationID].currentGoods =  stations[stationID].currentGoods + starttime * stations[stationID].Declinebymin;
                    stations[stationID].Locker = stations[stationID].Capacity - stations[stationID].currentGoods;
                    //commodity超過locker就要先塞回去不管後續的借還車= =凸操你媽的幹!!
                    if (commodity > stations[stationID].Locker)
                    {
                        //把多出的或塞回去改變卡車上的數量
                        truck.CurrentGoods += (commodity - (int)stations[stationID].Locker);
                        truck.CurrentSpace = truckcapacity - truck.CurrentGoods;
                        commodity = (int)stations[stationID].Locker;
                    }
                    if (stations[stationID].Locker < 0) { throw new Exception("big is here !!"); }
                    PnDTime = (commodity / PDVelocity);
                    if (starttime + PnDTime > timehorizon) { PnDTime = timehorizon - starttime; }
                    if (PDVelocity + stations[stationID].Declinebymin < 0 && stations[stationID].currentGoods + (PDVelocity + stations[stationID].Declinebymin) * PnDTime < 0)
                    {   //decline to 0 time
                        lackTime = -stations[stationID].currentGoods / (PDVelocity + stations[stationID].Declinebymin) + starttime;
                      
                        accumulatedtime = starttime + PnDTime - lackTime;
                        bikeAmount -= accumulatedtime * (PDVelocity + stations[stationID].Declinebymin);
                        stations[stationID].currentGoods = 0;
                        stations[stationID].Locker = stations[stationID].Capacity;
                        if (PnDTime < 0) throw new Exception("bug!!!");
                    }
                    else
                    {
                        if (stations[stationID].currentGoods + (PDVelocity + stations[stationID].Declinebymin) * PnDTime <= stations[stationID].Capacity)
                        {
                            stations[stationID].currentGoods += (PDVelocity + stations[stationID].Declinebymin) *PnDTime;
                        }
                        else //卡車送的超出站可以負荷的輛
                        {
                            //把多出的或塞回去改變卡車上的數量
                            truck.CurrentGoods += (commodity - (int)stations[stationID].Locker);
                            truck.CurrentSpace = truckcapacity - truck.CurrentGoods;
                            stations[stationID].currentGoods = stations[stationID].Capacity;
                            if (truck.CurrentGoods > truckcapacity || truck.CurrentSpace < 0)
                            {
                                throw new Exception();
                            }
                        }
                        stations[stationID].Locker = stations[stationID].Capacity - stations[stationID].currentGoods;
                    //    PnDTime = (TGS - truck.CurrentGoods) / PDVelocity;
                        if (PnDTime < 0) throw new Exception("bug!!!");
                    }
                    if (stations[stationID].Locker < 0) { throw new Exception("big is here !!"); }
                    lackTime = starttime + PnDTime-stations[stationID].currentGoods / stations[stationID].Declinebymin;
                    if (timehorizon > lackTime)
                    {
                        bikeAmount += stations[stationID].Declinebymin * (timehorizon - lackTime)*-1;
                        if (bikeAmount < 0) throw new Exception("bug!!!");
                    }
                }
                stations[stationID].currentGoods = Convert.ToInt32(initialGoods[stationID]);
                stations[stationID].Locker = stations[stationID].Capacity - stations[stationID].currentGoods;
                if (stations[stationID].Locker < 0) { throw new Exception("big is here !!"); }
                if (bikeAmount < 0) throw new Exception("bug!!!");
                return bikeAmount;
            }
            else if(stations[stationID].Surplusbymin > 0) //pickup bicycle station commodity must be negative
            {
                double surplustime;
                //卡車到達該站時該站的腳踏車就滿ㄌ
                if (stations[stationID].Locker/stations[stationID].Surplusbymin < starttime)
                {
                    surplustime = stations[stationID].Locker/stations[stationID].Surplusbymin;
                    accumulatedtime = starttime - surplustime;
                    bikeAmount += stations[stationID].Surplusbymin * accumulatedtime;
                    //commodity超過locker就要先塞回去不管後續的借還車= =凸操你媽的幹!!
                    if (-commodity > stations[stationID].Capacity)
                    {
                        //把多出的space塞回去改變卡車上的數量
                        truck.CurrentSpace -= (commodity + (int)stations[stationID].Capacity);
                        truck.CurrentGoods = truckcapacity - truck.CurrentSpace;
                        commodity = -(int)stations[stationID].Capacity;
                    }
                    if (bikeAmount < 0) throw new Exception("bug!!!");
                    PnDTime = -commodity / PDVelocity;
                    if (starttime + PnDTime > timehorizon) { PnDTime = timehorizon - starttime; }
                    if (-PDVelocity + stations[stationID].Surplusbymin > 0) //station increase rate > truck pickup velocity
                    {
                      
                        bikeAmount += (stations[stationID].Surplusbymin - PDVelocity) *PnDTime;
                        stations[stationID].currentGoods = stations[stationID].Capacity;
                        stations[stationID].Locker = 0;
                        if (PnDTime < 0) throw new Exception("bug!!!");
                    }
                    else
                    {
                        if (stations[stationID].Capacity >= (-PDVelocity + stations[stationID].Surplusbymin) * -PnDTime)
                        {
                            stations[stationID].Locker = (-PDVelocity + stations[stationID].Surplusbymin) * -PnDTime;
                        }
                        else
                        {
                            stations[stationID].Locker = stations[stationID].Capacity;
                            truck.CurrentSpace -= (commodity + (int)stations[stationID].Capacity);
                            truck.CurrentGoods = truckcapacity - truck.CurrentSpace;
                        }
                        if (stations[stationID].Locker < 0) { throw new Exception("big is here !!"); }
                        stations[stationID].currentGoods = stations[stationID].Capacity - stations[stationID].Locker;
                     //   PnDTime = (TGS - truck.CurrentSpace) / PDVelocity;
                        if (PnDTime < 0) throw new Exception("bug!!!");
                    }
                    if (commodity != 0)
                    {
                        surplustime = stations[stationID].Locker / stations[stationID].Surplusbymin + starttime+PnDTime;//又滿時
                        if (surplustime < timehorizon)
                        {
                            bikeAmount += stations[stationID].Surplusbymin * (timehorizon - surplustime);
                            if (bikeAmount < 0) throw new Exception("bug!!!");
                        }
                    }
                    else
                    {
                        if (surplustime < timehorizon)
                        {
                            bikeAmount = (timehorizon - surplustime) * stations[stationID].Surplusbymin;
                        }
                    }
                }
                else
                {
                    stations[stationID].currentGoods = stations[stationID].currentGoods + starttime * stations[stationID].Surplusbymin;
                    stations[stationID].Locker = stations[stationID].Capacity - stations[stationID].currentGoods;
                    //commodity超過locker就要先塞回去不管後續的借還車= =凸操你媽的幹!!
                    if (-commodity > stations[stationID].currentGoods)
                    {
                        //把多出的space塞回去改變卡車上的數量
                        truck.CurrentSpace -= (commodity + (int)stations[stationID].currentGoods);
                        truck.CurrentGoods = truckcapacity - truck.CurrentSpace;
                        commodity = -(int)stations[stationID].currentGoods;
                    }
                    PnDTime = -(commodity / PDVelocity);
                    if (starttime + PnDTime > timehorizon) { PnDTime = timehorizon - starttime; }
                    if (-PDVelocity + stations[stationID].Surplusbymin > 0 && stations[stationID].currentGoods + (-PDVelocity + stations[stationID].Surplusbymin) * -PnDTime >= stations[stationID].Capacity)
                    {   //increase to full time
                        surplustime = stations[stationID].Locker / (-PDVelocity + stations[stationID].Surplusbymin) + starttime;
                     
                        accumulatedtime = starttime + PnDTime - surplustime;
                        bikeAmount += accumulatedtime * (-PDVelocity + stations[stationID].Surplusbymin);
                        stations[stationID].currentGoods = stations[stationID].Capacity;
                        stations[stationID].Locker = 0;
                        if (PnDTime < 0)
                        {
                            throw new Exception("bug!!!");
                        }
                    }
                    else
                    {
                        if (stations[stationID].currentGoods - (-PDVelocity + stations[stationID].Surplusbymin) * -PnDTime >= 0)
                        {
                            stations[stationID].Locker = stations[stationID].Capacity - (stations[stationID].currentGoods - (-PDVelocity + stations[stationID].Surplusbymin) *-PnDTime);
                            if (stations[stationID].Locker < 0) { throw new Exception("big is here !!"); }
                        }
                        else //該站已經不能再拿走貨物ㄌ不然就會變成負的
                        {
                            truck.CurrentSpace -= (commodity + (int)stations[stationID].currentGoods);
                            truck.CurrentGoods = truckcapacity - truck.CurrentSpace;
                            stations[stationID].Locker = 0;
                            //改變卡車上的貨物數量
                            if (truck.CurrentGoods > truckcapacity || truck.CurrentSpace < 0)
                            {
                                throw new Exception();
                            }
                        }
                     //   PnDTime = (TGS - truck.CurrentSpace) / PDVelocity;
                        if (PnDTime < 0)
                        {
                            throw new Exception("bug!!!");
                        }
                    }
                    surplustime = stations[stationID].Locker / stations[stationID].Surplusbymin + starttime+PnDTime;
                    if (surplustime < timehorizon)
                    {
                        bikeAmount += stations[stationID].Surplusbymin * (timehorizon - surplustime);
                        if (bikeAmount < 0) throw new Exception("bug!!!");
                    }
                }

                stations[stationID].currentGoods = Convert.ToInt32(initialGoods[stationID]);
                stations[stationID].Locker = stations[stationID].Capacity - stations[stationID].currentGoods;
                if (stations[stationID].Locker < 0) { throw new Exception("bug is here !!"); }
                if (bikeAmount < 0) throw new Exception("bug!!!");
                return bikeAmount;
            }
            else //該站點無增減率時
            {
                if (commodity < 0)
                {
                    if(commodity+ stations[stationID].currentGoods<0)//無法拿的情況下
                    {

                        truck.CurrentSpace -= (commodity + (int)stations[stationID].currentGoods);
                        if (truck.CurrentGoods > truckcapacity || truck.CurrentGoods < 0 || truck.CurrentSpace > truckcapacity || truck.CurrentSpace < 0)
                        {
                            throw new Exception("bug");
                        }
                        truck.CurrentGoods = truckcapacity - truck.CurrentSpace;
                        if (truck.CurrentGoods > truckcapacity || truck.CurrentGoods < 0 || truck.CurrentSpace > truckcapacity || truck.CurrentSpace < 0)
                        {
                            throw new Exception("bug");
                        }
                    }
                }
                else if (commodity > 0)
                {
                    if (commodity > stations[stationID].Locker)//無法放的情況下
                    {
                        //把多出的或塞回去改變卡車上的數量
                        truck.CurrentGoods += (commodity - (int)stations[stationID].Locker);
                        if (truck.CurrentGoods > truckcapacity || truck.CurrentGoods < 0 || truck.CurrentSpace > truckcapacity || truck.CurrentSpace < 0)
                        {
                            throw new Exception("bug");
                        }
                        truck.CurrentSpace = truckcapacity - truck.CurrentGoods;
                        if (truck.CurrentGoods > truckcapacity || truck.CurrentGoods < 0 || truck.CurrentSpace > truckcapacity || truck.CurrentSpace < 0)
                        {
                            throw new Exception("bug");
                        }
                    }
                }
                return 0;
            }           
        }

        private void openfile(object sender, EventArgs e)
        {
            dataGridView.Rows.Clear();dataGridView.Columns.Clear();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                StreamReader sr = new StreamReader(openFileDialog.FileName);
                labelfilename.Text ="File: "+ openFileDialog.FileName;
                labeltitle.Text = "Title: "+ sr.ReadLine().ToString();
                totalStation = Convert.ToInt32( sr.ReadLine().ToString());
                timehorizon = Convert.ToDouble(sr.ReadLine());txbtimehorizon.Text = timehorizon.ToString();               
                truckcapacity = Convert.ToInt32(sr.ReadLine()); txbcapacity.Text = truckcapacity.ToString();
                speed = Convert.ToDouble(sr.ReadLine());txbspeed.Text = speed.ToString();
                PDVelocity = 1 / Convert.ToDouble(sr.ReadLine()); textBoxPnDTime.Text = (1 / PDVelocity).ToString();

                initialGoods = sr.ReadLine().Split(',');
                string[]capacity = sr.ReadLine().Split(',');
                string[] rate = sr.ReadLine().Split(',');
                string[] distance = sr.ReadLine().Split(',');
                string[] startTOother = sr.ReadLine().Split(',');
                string[] location = sr.ReadLine().Split(',');
                stations = new Station[totalStation];
                distanceMatrix = new double[totalStation, totalStation];
                startTOvertex = new double[totalStation];
                BUSA = new double[totalStation];BestUSA = new double[totalStation];
                for(int i = 0; i < stations.Length; i++)
                {
                      stations[i] = new Station(Convert.ToInt32(capacity[i]), Convert.ToInt32(initialGoods[i]), i, Convert.ToDouble(rate[i]));           
                }
                //for station location
                int c = 0;
                longitude = new double [totalStation+1];
                latitude = new double[totalStation+ 1];
                for(int a = 0; a < location.Length; a += 2)
                {
                    latitude[c] = Convert.ToDouble(location[a]);
                    longitude[c] = Convert.ToDouble(location[a + 1]);
                    c++;
                }
                c = 0;
                //目前仍不知道倉儲的位子
                for(int i = 0; i < totalStation; i++)
                {
                    for(int j = 0; j < totalStation; j++)
                    {
                        distanceMatrix[i, j] = Convert.ToDouble(distance[c]);
                        c++;
                    }
                }
                //加入從depot到其他vertex的距離
                for(int i = 0; i < totalStation; i++)
                {
                    startTOvertex[i] = Convert.ToDouble(startTOother[i]);
                }
                PandD = new int[totalStation];
                BestRouting = new int[totalStation];
                sr.Close();

                if (checkBox.Checked == true)
                {
                    for (int k = 0; k < totalStation; k++)
                    {
                        dataGridView.Columns.Add("station"+stations[k].StationID.ToString(),"station"+stations[k].StationID.ToString());
                        dataGridView.Columns[k].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        dataGridView.Columns[k].HeaderCell.ValueType = typeof(string);
                        dataGridView.Columns[k].HeaderCell.Value = "station" + stations[k].StationID.ToString();
                    }

                    dataGridView.Rows.Add();
                    dataGridView.Rows[0].HeaderCell.ValueType = typeof(string);
                    dataGridView.Rows[0].HeaderCell.Value = "cost";
                    dataGridView.Rows[1].HeaderCell.Value = "Best obj";
                }
            }           
            
        }

        private void getlowerbound(object sender, EventArgs e)
        {
            totalDistance = 0;
            Station[] ss = new Station[totalStation];
            //從0開始出發每次都找最短距離作為他的下一站
            stations.CopyTo(ss, 0);
            Array.Sort(ss);
            Array.Reverse(ss);
            Station[] Sequence = new Station[totalStation];
            int j = 0;
            for (int i = totalStation - 1; i > -1; i--)
            {
                Sequence[j] = ss[i]; Console.Write(Sequence[j].StationID + " "); j++;
                
            }
            double period = 0; //利用distancematrix除以車速即可得到現在的時間
            double cost = 0.0;
            period = (startTOvertex[Sequence[0].StationID] / speed + (truckcapacity/2)/PDVelocity);//抵達第一站的時間 
            totalDistance += startTOvertex[Sequence[0].StationID]; //抵達第一站的距離
            int a = 0;Truck truck = new Truck(truckcapacity, truckcapacity/2);
            int[] temp = new int[totalStation];
            while(period <= timehorizon && a < totalStation - 1)
            {

                if (Sequence[a].Declinebymin < 0)
                {
                    temp[a] = truck.CurrentGoods;
                    cost += HoldingShortageSurplus(period, Sequence[a].StationID, truck.PickupDelivery(truckcapacity / 2), ref truck, temp[a]);
                    temp[a] -= truck.CurrentGoods;
                }
                else if (Sequence[a].Surplusbymin > 0)
                {
                    temp[a] = truck.CurrentSpace;
                    cost += HoldingShortageSurplus(period, Sequence[a].StationID, truck.PickupDelivery(-truckcapacity / 2), ref truck, temp[a]);
                    temp[a] = truck.CurrentSpace - temp[a];
                }
                period += (distanceMatrix[Sequence[a].StationID, Sequence[a + 1].StationID] /speed+ Math.Abs(temp[a] / PDVelocity));//到下一站的時間點
                totalDistance += distanceMatrix[Sequence[a].StationID, Sequence[a + 1].StationID];
                a++;
            }
            if(period > timehorizon) { totalDistance-= distanceMatrix[Sequence[a-1].StationID, Sequence[a].StationID]; }
            //因為上面最後一次部會做所以...
            if (Sequence[a].Declinebymin < 0 && period <= timehorizon)
            {
                temp[a] = truck.CurrentGoods;
                cost += HoldingShortageSurplus(period, Sequence[a].StationID, truck.PickupDelivery(truckcapacity / 2), ref truck, temp[a]);
                temp[a] -= truck.CurrentGoods;
            }
            else if(Sequence[a].Surplusbymin > 0 && period <= timehorizon)
            {
                temp[a] = truck.CurrentSpace;
                cost += HoldingShortageSurplus(period, Sequence[a].StationID, truck.PickupDelivery(-truckcapacity / 2), ref truck, temp[a]);
                temp[a] = truck.CurrentSpace - temp[a];
            }
            else { a--; }
           
            //全部跑完要回去depot
            if (a == totalStation - 1)
            {
                totalDistance += startTOvertex[Sequence[a].StationID];
            }
            else
            {
                //沒有跑完的車站繼續算成本
                for (int b = a; b < totalStation - 1; b++)
                {
                    truck.CurrentGoods = 0; truck.CurrentSpace = truckcapacity;
                    if (stations[Sequence[b].StationID].Surplusbymin > 0)
                    {
                        cost += HoldingShortageSurplus(timehorizon, Sequence[b].StationID, 0, ref truck, truckcapacity);
                    }
                    else
                    {
                        cost += HoldingShortageSurplus(timehorizon, Sequence[b].StationID, 0, ref truck, 0);
                    }
                   
                }
            }
            int totalpnd = 0;
            richTextBox.Text = "routing path is depot ";
            for(int i = 0; i < a+1; i++)
            {
                richTextBox.Text += (Sequence[i].StationID + " "); 
            }
            richTextBox.Text += ("\n " + "pickup and delivery amount : ");
            for (int i = 0; i <= a; i++)
            {
                richTextBox.Text += (temp[i] + " ");
                totalpnd += Math.Abs(temp[i]);
            }
            totalpnd += truckcapacity / 2;
            richTextBox.Text +=("\n" +"********performance********");
            richTextBox.Text +=  ("\n" + "total unsatisfied amount is :" + cost);
            richTextBox.Text += ("\n" + "total distance is :" + totalDistance );
            richTextBox.Text += ("\n" + "total pickup & delivery: " + totalpnd);
        }

        private void nearestInsert(object sender, EventArgs e)
        {
            totalDistance = 0;
            List<int> seq = new List<int>();
            //List<List<int>> ss = new List<List<int>>();
            double min = double.MaxValue; int last = -1;int next = -1;
            double[,] d = new double[totalStation, totalStation];

            for (int a = 0; a < totalStation; a++)
            {
                for (int b = 0; b < totalStation; b++)
                {
                    d[a, b] = distanceMatrix[a, b];
                }
            }
            for (int i = 1; i < totalStation; i++)
            {
                if (min > startTOvertex[i]) { min = startTOvertex[i]; last = i; }
            }
            //min = double.MaxValue;int idx = -1;
            //for(int a = 0;a < stations.Length; a++)
            //{
            //    totalDistance = 0;
            //    ss.Add(Greedy(a));
            //    for(int b = 0; b < stations.Length-1; b++)
            //    {
            //        totalDistance += distanceMatrix[ss.ElementAt(a).ElementAt(b), ss.ElementAt(a).ElementAt(b + 1)];
            //    }
            //    if (min > totalDistance) { min = totalDistance; idx = a; }
            //}

            seq.Add(last);
            do
            {
                min = double.MaxValue;
                for (int i = 0; i < totalStation; i++)
                {
                    if (d[last, i] != 0 && min > d[last, i])
                    {
                        min = d[last, i]; next = i;
                    }
                }
                for (int j = 0; j < totalStation; j++)
                {
                    d[j, last] = 0;
                }
                seq.Add(next);
                last = next;
            } while (check(d, next));
            int[] routing = new int[seq.Count];
            seq.CopyTo(routing);
            //int[] routing = new int[ss.ElementAt(idx).Count];
            // ss.ElementAt(idx).CopyTo(routing);
            //排成可繞之形成
            //for(int a = 0; a < stations.Length; a++)
            //{
            //    if(ss.ElementAt(idx).ElementAt(a) == last)
            //    {
            //        last = a;
            //        break;
            //    }
            //}
            //for (int a = last; a < stations.Length; a++)
            //{
            //    routing[a-last] = ss.ElementAt(idx).ElementAt(a);
            //}
            //for(int b = 0;b < last; b++)
            //{
            //    routing[stations.Length-last + b] = ss.ElementAt(idx).ElementAt(b);
            //}
            //totalDistance = 0;
            double period = 0; //利用distancematrix除以車速即可得到現在的時間
            double cost = 0.0;
            string dist = null;
            Truck truck = new Truck(truckcapacity,truckcapacity/2);
            period = (startTOvertex[routing[0]] / speed )+ (truckcapacity/2)/PDVelocity;//抵達第一站的時間 
            totalDistance += startTOvertex[routing[0]]; //抵達第一站的距離  
            dist+= (startTOvertex[routing[0]].ToString()+"->");
            int c = 0;int[] temp = new int[totalStation];
            while (c < totalStation-1 && period <= timehorizon)
            {
                if (stations[routing[c]].Declinebymin < 0)
                {
                    temp[c] = truck.CurrentGoods;
                    cost += HoldingShortageSurplus(period, routing[c],truck.PickupDelivery( truckcapacity / 2), ref truck, temp[c]);
                    temp[c] -= truck.CurrentGoods;
                }
                else if (stations[routing[c]].Surplusbymin > 0)
                {
                    temp[c] = truck.CurrentSpace;
                    cost += HoldingShortageSurplus(period, routing[c], truck.PickupDelivery(-truckcapacity / 2), ref truck, temp[c]);
                    temp[c] = truck.CurrentSpace-temp[c];
                }
                else { temp[c] = 0; }
                period += (distanceMatrix[routing[c], routing[c + 1]]/speed+ Math.Abs(temp[c] / PDVelocity));//到下一站的時間點
                totalDistance += distanceMatrix[routing[c], routing[c + 1]];
                dist += (distanceMatrix[routing[c], routing[c + 1]] + "->");
                c++;
            }
            //因為上面最後一次部會做所以...
            if (stations[routing[c]].Declinebymin < 0 && period <= timehorizon)
            {
                temp[c] = truck.CurrentGoods;
                cost += HoldingShortageSurplus(period, routing[c], truck.PickupDelivery(truckcapacity / 2), ref truck, temp[c]);
                temp[c] -= truck.CurrentGoods;
            }
            else if(stations[routing[c]].Surplusbymin > 0 && period <= timehorizon)
            {
                temp[c] = truck.CurrentSpace;
                cost += HoldingShortageSurplus(period, routing[c], truck.PickupDelivery(-truckcapacity / 2), ref truck, temp[c]);
                temp[c] = truck.CurrentSpace - temp[c];
            }
            else { c--; }
            
            //全部跑完要回去depot
            if (c == totalStation - 1)
            {
                totalDistance += startTOvertex[routing[c]];
                dist += (startTOvertex[routing[c]].ToString() + "->");
            }
            else
            {
                //沒有跑完的車站繼續算成本
                truck.CurrentGoods = 0;truck.CurrentSpace = truckcapacity;
                for (int b = c; b < totalStation - 1; b++)
                {
                    if (stations[routing[b]].Surplusbymin > 0)
                    {
                        cost += HoldingShortageSurplus(timehorizon, routing[b], 0, ref truck, truckcapacity);
                    }
                    else
                    {
                        cost += HoldingShortageSurplus(timehorizon, routing[b], 0, ref truck, 0);
                    }
                }
            }
            richTextBox.Text = "routing path is: depot ";
            for (int i = 0; i < c+1; i++)
            {
                richTextBox.Text += (routing[i] + " ");
            }
            int tpnd = 0;
            richTextBox.Text += "\n" + "pickup and delivery ";
            for (int i = 0; i < c + 1; i++)
            {
                richTextBox.Text += (temp[i] + " ");
                tpnd += Math.Abs( temp[i] );
            }
            tpnd += truckcapacity / 2;
            richTextBox.Text += "\n" + "********performance********";
            richTextBox.Text += ("\n" + "total unsatisfied amount is :" + cost);
            richTextBox.Text += ("\n" + "total distance is " + totalDistance);
            richTextBox.Text += ("\n" + "total pickup & delivery is " + tpnd);
            richTextBox.Text += ("\n" + "improvement rate: " + Math.Round(1-cost / NotDoAnything, 2).ToString());
            richTextBox.Text += ("\n" + dist);
        }
        private bool check(double[,] m,int n)
        {  bool flag = false;
            for (int i = 0; i < totalStation; i++)
            {
                if (m[n, i] != 0) { flag = true; break; }
            }
            return flag;
        }

        private List<int> Greedy(int last)
        {
            List<int> seq = new List<int>();
            double min = double.MaxValue; int next = -1;
            double[,] d = new double[totalStation, totalStation];
            //seq.Add(0); //must begin at 0
            for (int a = 0; a < totalStation; a++)
            {
                for (int b = 0; b < totalStation; b++)
                {
                    d[a, b] = distanceMatrix[a, b];
                }
            }
            seq.Add(last);
            do
            {
                min = double.MaxValue;
                for (int i = 0; i < totalStation; i++)
                {
                    if (d[last, i] != 0 && min > d[last, i])
                    {
                        min = d[last, i]; next = i;
                    }
                }
                for (int j = 0; j < totalStation; j++)
                {
                    d[j, last] = 0;
                }
                seq.Add(next);
                last = next;
            } while (check(d, next));
            return seq;
        }

        private void NBS(object sender, EventArgs e)
        {
            //int[] ans = new int[] {62, 110, 244, 52, 46, 99, 71, 58, 75, 98, 78, 217, 161, 34, 33, 51, 235, 172, 120, 252, 174, 226, 239, 184, 48, 84, 86, 37, 131, 44, 60, 65, 103, 97, 93, 68,0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,35,36,38,39,40,41,42,43,45,47,49,50,53,54,55,56,57,59,61,63,64,66,67,69,70,72,73,74,76,77,79,80,81,82,83,85,87,88,89,90,91,92,94,95,96,100,101,102,104,105,106,107,108,109,111,112,113,114,115,116,117,118,119,121,122,123,124,125,126,127,128,129,130,132,133,134,135,136,137,138,139,140,141,142,143,144,145,146,147,148,149,150,151,152,153,154,155,156,157,158,159,160,162,163,164,165,166,167,168,169,170,171,173,175,176,177,178,179,180,181,182,183,185,186,187,188,189,190,191,192,193,194,195,196,197,198,199,200,201,202,203,204,205,206,207,208,209,210,211,212,213,214,215,216,218,219,220,221,222,223,224,225,227,228,229,230,231,232,233,234,236,237,238,240,241,242,243,245,246,247,248,249,250,251,253,254,255,
            //   0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,-14,0,0,0,0,0,0,0,0,0,-8,0,29,0,27,0,0,-18,2,0,0,0,0,0,-27,0,-12,0,4,0,0,4,0,0,-21,0,0,0,0,0,0,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,15,0,0,0,-18,25,1,0,0,0,16,0,0,0,0,0,0,-11,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,-8,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,30,0,0,0,0,0,0,0,0,0,0,0,5,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,-8,0,0,0,0,0,0,0,0,1,0,0,0,-24,0,0,0,0,-9,0,0,0,0,0,0,0,0,0,0,0
            //    ,16,0
            //};
            //initial good , start time
            int[] ans = new int[]
            {137, 236, 15, 12, 149, 87, 8, 18, 9, 11, 162, 220, 29, 246, 125, 31, 242, 74, 70, 10, 186, 1, 248, 32, 230, 171, 254, 173, 17, 54, 51, 235, 172, 33, 34, 59, 50, 13, 85, 240, 175, 120, 252, 174, 226, 239, 184, 48, 62, 110, 244, 52, 46, 100, 99, 71, 193, 57, 114, 159, 98, 78,0,2,3,4,5,6,7,14,16,19,20,21,22,23,24,25,26,27,28,30,35,36,37,38,39,40,41,42,43,44,45,47,49,53,55,56,58,60,61,63,64,65,66,67,68,69,72,73,75,76,77,79,80,81,82,83,84,86,88,89,90,91,92,93,94,95,96,97,101,102,103,104,105,106,107,108,109,111,112,113,115,116,117,118,119,121,122,123,124,126,127,128,129,130,131,132,133,134,135,136,138,139,140,141,142,143,144,145,146,147,148,150,151,152,153,154,155,156,157,158,160,161,163,164,165,166,167,168,169,170,176,177,178,179,180,181,182,183,185,187,188,189,190,191,192,194,195,196,197,198,199,200,201,202,203,204,205,206,207,208,209,210,211,212,213,214,215,216,217,218,219,221,222,223,224,225,227,228,229,231,232,233,234,237,238,241,243,245,247,249,250,251,253,255,
            0,5,0,0,0,0,0,0,14,-2,-6,-5,-2,0,0,-10,0,-8,-5,0,0,0,0,0,0,0,0,0,0,0,0,-6,5,0,-12,0,0,0,0,0,0,0,0,0,0,0,2,0,12,0,0,-1,12,0,-5,0,0,-4,0,-2,0,0,0,0,0,0,0,0,0,0,2,0,0,0,-8,0,0,0,14,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,-12,0,0,0,-2,0,0,0,0,0,0,0,0,0,0,-8,0,0,0,0,0,0,0,0,0,0,0,14,0,0,0,0,0,0,0,0,0,0,0,-2,0,0,0,0,0,0,0,0,0,0,0,0,-2,0,0,0,0,0,0,0,0,14,12,0,7,0,0,0,0,0,0,0,0,0,2,0,0,0,0,0,0,0,0,-8,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,14,0,0,0,0,0,-4,0,0,0,-5,0,0,0,0,2,0,0,0,-10,0,0,14,0,-2,0,0,0,-7,0,0,0,7,0,0,0,
            14,0           };

            richTextBox.Text += "\n" + "********performance********";
            richTextBox.Text += "\n" + "unsatisfied amount :" + ComputeTimeAverageHoldingShortageLevel(ans);
            richTextBox.Text += ("\n" + "total distance is " + totalDistance);
            //richTextBox.Text += "\n" + "Total pickup & delivery : " + totalPandD;

        }

        private void saveToExcel(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWB;
            Excel.Worksheet excelWS;
            //Excel.Range oRng;
            saveFileDialog.Filter = "Excel活頁簿(*.xlsx)|*.xlsx";
            if (save.ShowDialog() == DialogResult.OK)
            {
                string fileName = save.FileName;
                excelApp.Workbooks.Add();
                excelWB = excelApp.Workbooks[1];
                excelWS = excelWB.Worksheets[1];

                excelWS.Cells[1, 1] = "truck start time";
                excelWS.Cells[2, 1] = "truck initial goods";
                excelWS.Cells[3, 1] = "routing sequence";
                excelWS.Cells[4, 1] = "P&D amount";
                excelWS.Cells[5, 1] = "withoutaction";

                excelWS.Cells[7, 1] = "total distance";
                excelWS.Cells[8, 1] = "total P&D";
                excelWS.Cells[9, 1] = "unsatisfied amount";
                excelWS.Cells[10, 1] = "improvement rate";

               

                excelWS.Cells[1, 2] = pga.SoFarTheBestSolution[totalStation*2].ToString("0.00");
                excelWS.Cells[2, 2] = pga.SoFarTheBestSolution[totalStation * 2 + 1].ToString("0.00");
                string r = "", pd = "";double totalPandD = 0.0;
                for (int i = 0; i < soFarfinishStation + 1; i++)
                {
                    r += BestRouting[i].ToString() + ",";
                    pd += PandD[i] + ",";
                    //excelWS.Cells[3, 2 + i] = BestRouting[i];
                    //excelWS.Cells[4, 2 + i] = PandD[i]; 
                    totalPandD += Math.Abs(PandD[i]);
                }
                excelWS.Cells[3, 2] = r;
                excelWS.Cells[4, 2] = pd;
                excelWS.Cells[5, 2] = NotDoAnything.ToString("0.00");
                excelWS.Cells[7, 2] = totalDistance.ToString("0.00");
                excelWS.Cells[8, 2] = (totalPandD+ pga.SoFarTheBestSolution[totalStation * 2]).ToString();
                excelWS.Cells[9, 2] = Math.Round(pga.SoFarTheBestObjective, 2).ToString();
                excelWS.Cells[10, 2] = Math.Round(1 - pga.SoFarTheBestObjective/NotDoAnything, 2).ToString("0.00");

                excelWB.SaveAs(fileName);
                excelWB.Close();
            }

            excelApp.Quit();
            //釋放Excel資源
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            GC.Collect();
        }

        private void DrwaRoutes(object sender, PaintEventArgs e)
        {
            if (stations != null)
            {
                Station[] ss = new Station[totalStation];
                //從0開始出發每次都找最短距離作為他的下一站
                stations.CopyTo(ss,0);
                Array.Sort(ss);
                Array.Reverse(ss);
                int [] s = new int[Convert.ToInt32( totalStation * 0.2)];
                for(int i = 0; i < s.Length; i++)
                {
                    s[i] = ss[i].StationID;
                }
                
                w = e.ClipRectangle.Width;
                h = e.ClipRectangle.Height;

                xmin = longitude.Min();
                W = longitude.Max() - xmin;  //最大經度與最小經度的差

                ymax = latitude.Max();
                H = ymax - latitude.Min();  //最大緯度與最小緯度的差

                scale = w / W * 0.8;   //算出width的scale
                scale2 = h / H * 0.8;  //算出hight的scale

                xmiddle = longitude[totalStation / 2];   //尚待修正
                ymiddle = latitude[17];    //尚待修正代表站點17的latitude當作中間值

                Rectangle rec = new Rectangle(0, 0, 10, 10);  //先初始化，x與y之後使用者自訂 裡面的數字即是圓的大小
                Rectangle rec2 = new Rectangle(0, 0, 0, 0);
                //繪製站點位置與名稱於UI
                for (int i = 0; i < longitude.Length; i++)
                {
                    if (i == 0)
                    {
                        //繪製站點位置圓圈線段
                        rec.X = getx(longitude[i]);
                        rec.Y = gety(latitude[i]);

                        e.Graphics.FillRectangle(Brushes.Red, rec);
                        //繪製站點名稱字樣
                        SizeF sz = e.Graphics.MeasureString("Depot", myfont);
                        rec2.Width = (int)sz.Width;
                        rec2.Height = (int)sz.Height;

                        rec2.X = (int)(getx(longitude[i]) - (rec2.Width) * 0.33);
                        rec2.Y = gety(latitude[i]) + 7;

                        e.Graphics.DrawString("Depot", myfont, Brushes.Red, rec2.Location);
                    }
                    else
                    {
                        //繪製站點位置圓圈線段
                        rec.X = getx(longitude[i]);
                        rec.Y = gety(latitude[i]);
                        if(stations[i-1].Rate < 0)
                        {
                            e.Graphics.DrawEllipse(Pens.Green, rec);
                            if(s.Contains(stations[i- 1].StationID))
                            {
                                e.Graphics.FillEllipse(Brushes.Green, rec);
                            }
                        }
                        else if(stations[i-1].Rate > 0)
                        {
                            e.Graphics.DrawEllipse(Pens.Blue, rec);
                            if (s.Contains( stations[i - 1].StationID))
                            {
                                e.Graphics.FillEllipse(Brushes.Blue, rec);
                            }
                        }
                        else
                        {
                            e.Graphics.DrawEllipse(Pens.Brown, rec);
                        }

                        //繪製站點名稱字樣
                        SizeF sz = e.Graphics.MeasureString(stations[i-1].StationID.ToString(), myfont);
                        rec2.Width = (int)sz.Width;
                        rec2.Height = (int)sz.Height;

                        rec2.X = (int)(getx(longitude[i]) - (rec2.Width) * 0.2);
                        rec2.Y = gety(latitude[i]) + 8;

                        e.Graphics.DrawString(stations[i-1].StationID.ToString(), myfont, Brushes.Black, rec2.Location);
                    }

                }

                mypen.Color = Color.FromArgb(0, 0, 0);

                //繪製卡車移動路徑
                e.Graphics.DrawLine(mypen, getx(longitude[0]) + 7, gety(latitude[0]) + 7, getx(longitude[BestRouting[0]+1]) + 7, gety(latitude[BestRouting[0]+1]) + 7);  //畫出depot到第一站的路徑
                for (int i = 0; i < soFarfinishStation; i++)
                {
                    e.Graphics.DrawLine(mypen, getx(longitude[BestRouting[i]+1]) + 7, gety(latitude[BestRouting[i]+1]) + 7, 
                        getx(longitude[BestRouting[i + 1]+1]) + 7, gety(latitude[BestRouting[i + 1]+1]) + 7);  //畫出站與站間的路徑
                    e.Graphics.DrawLine(mypen, getx(longitude[BestRouting[i]+1]) + 7, gety(latitude[BestRouting[i]+1]) + 7, 
                        getx(longitude[BestRouting[i + 1]+1]) + 7, gety(latitude[BestRouting[i + 1]+1]) + 7);  //畫出最後一站的路徑
                                                                               
                }
            }
        }

        int getx(double x)
        {
            int temp = 0;
            if (x == xmiddle)
                temp = w / 2;
            else if (x < xmiddle)
                temp = (int)(w / 2 - (xmiddle - x) * scale);
            else
                temp = (int)(w / 2 + (x - xmiddle) * scale);

            return temp;
        }

        int gety(double y)
        {
            int temp = 0;
            if (y == ymiddle)
                temp = (int)(h * 0.5);
            else if (y > ymiddle) //若緯度較高，則繪製較上方位置(y減少)
                temp = (int)(h * 0.5 - (y - ymiddle) * scale2);
            else                  //若緯度較低，則繪製較下方位置(y增加)
                temp = (int)(h * 0.5 + (ymiddle - y) * scale2);

            return temp;
        }
    }
}
