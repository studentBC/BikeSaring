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

namespace BikeSharingSystem
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }
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
        int truckcapacity;
        int totalStation;
        int finishStation;
        int soFarfinishStation; //so far the best solution complete station amount
        Station[] stations;
        private void reset(object sender, EventArgs e)
        {
            //pga = new PermutationGA(totalStation*2+2, OptimizationType.Min,ComputeTimeAverageHoldingShortageLevel,stations);
            //timehorizon = Convert.ToDouble(txbtimehorizon.Text.ToString());
            speed = Convert.ToDouble(txbspeed.Text.ToString());
            truckcapacity = pga.TruckCapacity;
            //truckcapacity = Convert.ToInt32(txbcapacity.Text.ToString());
            //pga.TruckCapacity = truckcapacity;
            //pga.TimeHorizon = timehorizon;
            if (checkBoxGreedy.Checked == true)
            {
                pga.Greedyinitialize(distanceMatrix,startTOvertex);
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
                NotDoAnything += HoldingShortageSurplus(timehorizon, i, 0, ref uselesstruck);
            }
            //update data grid view
            //dataGridView.Rows.Clear();
            //dataGridView.Columns.Clear();

            //for (int m = 0; m < this.numberOfJobs; m = c + 1)
            //{
            //    dataGridView1.Rows[l].Cells[m].Value = this.setupTimes[l, m];
            //    c = m;
            //}


        }
     
        private void createPGA(object sender, EventArgs e)
        {
            butreset.Enabled = true;
            butrun1.Enabled = true;
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
         
            double unsatisfiedAmount = pga.SoFarTheBestObjective;
            labelBO.Text = "BestObjective  " + Convert.ToString(unsatisfiedAmount);
            labelBO.Text.ToString();
            richTextBox.Text = "繞站順序: depot " ;
            for(int i = 0; i < soFarfinishStation+1; i++)
            {
                //我忘記解碼fuck fuck fuck fuck fuck fuck !! 
                //   richTextBox.Text += pga.SoFarTheBestSolution[i].ToString()+" ";             
                richTextBox.Text += BestRouting[i].ToString() + " ";
            }
         //   int[] a = Decode(pga.SoFarTheBestSolution); debug用
            if (soFarfinishStation == totalStation - 1)
            {
                richTextBox.Text += "depot";
            }
            double totalPandD = 0;
            richTextBox.Text +="\n"+"pickup & delivery amount: ";
            for(int i = 0; i <= soFarfinishStation; i++)
            {
                richTextBox.Text += PandD[i].ToString() + "  ";
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
            if(totalDistance >= speed*timehorizon) { throw new Exception("bug!!!"); }
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
            richTextBox.Text += "\n" + "improvement rate :" + Math.Round((unsatisfiedAmount / NotDoAnything),2)+"\n";
            //richTextBox.AppendText("routing distance "+dist);

           
            labelMax.Text ="Max :" + NotDoAnything.ToString();
            chart.Series[3].Points.AddXY((double)pga.IterationCount, NotDoAnything);

            // chart.Update();
            //讓圖形跑超過50代次時可跑的較緩慢
            //if ((pga.IterationCount > 50) && (!pga.terminationConditionMet())) { Cursor = System.Windows.Forms.Cursors.WaitCursor; }
            //else { Cursor = System.Windows.Forms.Cursors.Default; }
        }

        private void runToEnd(object sender, EventArgs e)
        {
            for (int i = 0; i < pga.IterationLimit; i++)
            {
                oneIteration(null, null);
            }
            ppg.Update();
        }
        double min = Double.MaxValue;
        
        private double ComputeTimeAverageHoldingShortageLevel(int[] chromosomes)
        {

            Truck truck = new Truck(truckcapacity, chromosomes[totalStation * 2]);
            //period為卡車起始出發時間
            double USA = 0; //USA = unsatisfied amount
            double period = chromosomes[totalStation * 2 + 1]; //利用distancematrix除以車速即可得到現在的時間
            double totalUSA = NotDoAnything; 
            totalDistance = 0;
            int[] routingSequence = Decode(chromosomes);
            period += (startTOvertex[routingSequence[0]] / speed);//抵達第一站的時間 
       //   totalDistance += startTOvertex[routingSequence[0]]; //抵達第一站的距離
            int i = 0; int[] temp = new int[totalStation];
            while (period <= timehorizon && i < totalStation - 1)
            {
                if(i == 0) { totalUSA = 0; }
                temp[i] = truck.PickupDelivery(chromosomes[routingSequence[i] + totalStation]);
                USA = HoldingShortageSurplus(period, routingSequence[i], temp[i],ref truck);
                totalUSA += USA;
                period += ((distanceMatrix[routingSequence[i], routingSequence[i + 1]] / speed ));//到下一站的時間點
                //totalDistance += distanceMatrix[routingSequence[i], routingSequence[i + 1]];
                i++;
                //update datagrid view for debugging
                if(checkBox.Checked == true)
                {
                    dataGridView.Rows[0].Cells[routingSequence[i]].Value = USA.ToString();
                }
            }
            if (i != 0 && period <= timehorizon)
            {
                //因為上面最後一次不會做到所以...
                temp[i] = truck.PickupDelivery(chromosomes[routingSequence[i] + totalStation]);
                USA = HoldingShortageSurplus(period, routingSequence[i], temp[i], ref truck);
                totalUSA += USA;
                finishStation = i;
            }else if(i != 0) { finishStation = i-1; }
            
   
            if (finishStation != totalStation - 1)
            {
                //沒有跑完的車站繼續算成本
                for (int a = finishStation; a < totalStation ; a++)
                {
                    USA= HoldingShortageSurplus(timehorizon, routingSequence[a], 0,ref truck);
                    totalUSA += USA;
                    if (checkBox.Checked == true)
                    {
                        dataGridView.Rows[0].Cells[routingSequence[a]].Value = USA.ToString();
                    }
                }
            }
            if(min > totalUSA)
            {
                min = totalUSA;
                soFarfinishStation = finishStation;
                temp.CopyTo(PandD,0);
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
        private double HoldingShortageSurplus(double starttime,int stationID ,int commodity,ref Truck truck)
        {
            double bikeAmount = 0;
            double accumulatedtime = 0;
          //  Truck t = truck;
            if (stations[stationID].Declinebymin < 0) //deliver bicycle to the station commodity must be positive
            {
                double lackTime = 0.0;
                //卡車到達該站時該站的腳踏車就沒ㄌ
                if ( stations[stationID].currentGoods / stations[stationID].Declinebymin * -1 < starttime)
                {
                    lackTime = stations[stationID].currentGoods / stations[stationID].Declinebymin * -1;
                    accumulatedtime = starttime - lackTime;
                    bikeAmount += stations[stationID].Declinebymin * accumulatedtime*-1;
                    if (bikeAmount < 0) throw new Exception("bug!!!");
                    if (stations[stationID].Capacity >= commodity)//如果給的數量超過站的容量就G囉
                    {
                        stations[stationID].currentGoods = commodity;
                    }
                    else
                    {
                        stations[stationID].currentGoods = stations[stationID].Capacity;
                        truck.CurrentGoods += (commodity - (int)stations[stationID].Capacity);
                        truck.CurrentSpace = truckcapacity - truck.CurrentGoods;
                    }
                    stations[stationID].Locker = stations[stationID].Capacity - stations[stationID].currentGoods;
                    if(stations[stationID].Locker < 0) { throw new Exception("big is here !!"); }
                    if (commodity != 0)
                    {
                        lackTime = stations[stationID].currentGoods / stations[stationID].Declinebymin * -1 + starttime;//又缺為0時
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
                    stations[stationID].currentGoods = Convert.ToInt32( stations[stationID].currentGoods + starttime * stations[stationID].Declinebymin);
                    stations[stationID].Locker = stations[stationID].Capacity - stations[stationID].currentGoods;
                    if (stations[stationID].Locker < 0) { throw new Exception("big is here !!"); }
                    if (stations[stationID].currentGoods + commodity <= stations[stationID].Capacity)
                    {
                        stations[stationID].currentGoods += (commodity);
                    }
                    else //卡車送的超出站可以負荷的輛
                    {
                        //把多出的或塞回去改變卡車上的數量
                        truck.CurrentGoods += (int)(commodity - stations[stationID].Locker);
                        truck.CurrentSpace = truckcapacity - truck.CurrentGoods;
                        stations[stationID].currentGoods = stations[stationID].Capacity;
                        if (truck.CurrentGoods > truckcapacity || truck.CurrentSpace < 0)
                        {
                            throw new Exception();
                        }
                    }
                    stations[stationID].Locker = stations[stationID].Capacity - stations[stationID].currentGoods;
                    if (stations[stationID].Locker < 0) { throw new Exception("big is here !!"); }
                    lackTime = stations[stationID].currentGoods / stations[stationID].Declinebymin*-1 + starttime;
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
                    if (bikeAmount < 0) throw new Exception("bug!!!");
                    if (stations[stationID].Capacity >= -commodity)
                    {
                        stations[stationID].Locker = -1 * commodity;
                    }
                    else
                    {
                        stations[stationID].Locker = stations[stationID].Capacity;
                        truck.CurrentSpace -= (int)(commodity + stations[stationID].Capacity);
                        truck.CurrentGoods = truckcapacity - truck.CurrentSpace;
                    }
                    if (stations[stationID].Locker < 0) { throw new Exception("big is here !!"); }
                    if (commodity != 0)
                    {
                        stations[stationID].currentGoods = stations[stationID].Capacity - stations[stationID].Locker;
                        surplustime = stations[stationID].Locker / stations[stationID].Surplusbymin + starttime;//又滿時
                        if (surplustime < timehorizon)
                        {
                            bikeAmount += stations[stationID].Surplusbymin * (timehorizon - surplustime);
                            if (bikeAmount < 0) throw new Exception("bug!!!");
                        }
                    }
                    else
                    {
                        stations[stationID].currentGoods = stations[stationID].Capacity - stations[stationID].Locker;
                        if (surplustime < timehorizon)
                        {
                            bikeAmount = (timehorizon - surplustime) * stations[stationID].Surplusbymin;
                        }
                    }
                }
                else
                {
                    stations[stationID].currentGoods = Convert.ToInt32(stations[stationID].currentGoods + starttime * stations[stationID].Surplusbymin);
                    if (stations[stationID].currentGoods + commodity >= 0)
                    {
                        stations[stationID].Locker = stations[stationID].Capacity - (stations[stationID].currentGoods + commodity);
                        if (stations[stationID].Locker < 0) { throw new Exception("big is here !!"); }
                    }
                    else //該站已經不能再拿走貨物ㄌ不然就會變成負的
                    {
                        truck.CurrentSpace -= (int)(commodity + stations[stationID].currentGoods);
                        truck.CurrentGoods = truckcapacity - truck.CurrentSpace;
                        stations[stationID].Locker = 0;
                        //改變卡車上的貨物數量
                        if (truck.CurrentGoods > truckcapacity || truck.CurrentSpace < 0)
                        {
                            throw new Exception();
                        }
                    }
                    surplustime = stations[stationID].Locker / stations[stationID].Surplusbymin + starttime;
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
            return 0;
        }

        private void openfile(object sender, EventArgs e)
        {
            dataGridView.Rows.Clear();dataGridView.Columns.Clear();
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                StreamReader sr = new StreamReader(openFileDialog.FileName);
                totalStation = Convert.ToInt32( sr.ReadLine().ToString());
                timehorizon = Convert.ToDouble(sr.ReadLine());txbtimehorizon.Text = timehorizon.ToString();
                truckcapacity = Convert.ToInt32(sr.ReadLine()); txbcapacity.Text = truckcapacity.ToString();
                speed = Convert.ToDouble(sr.ReadLine());txbspeed.Text = speed.ToString();
                
                initialGoods = sr.ReadLine().Split(',');
                string[]capacity = sr.ReadLine().Split(',');
                string[] rate = sr.ReadLine().Split(',');
                string[] distance = sr.ReadLine().Split(',');
                string[] startTOother = sr.ReadLine().Split(',');
                stations = new Station[totalStation];
                distanceMatrix = new double[totalStation, totalStation];
                startTOvertex = new double[totalStation];
                
                for(int i = 0; i < stations.Length; i++)
                {
                      stations[i] = new Station(Convert.ToInt32(capacity[i]), Convert.ToInt32(initialGoods[i]), i, Convert.ToDouble(rate[i]));             
                }
                //for benchmark
               // stations[1].currentGoods = 40;stations[1].currentSpace = 0;
                int c = 0;
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
            Station[] Sequence = new Station[totalStation];
            int j = 0;
            for (int i = totalStation - 1; i > -1; i--)
            {
                Sequence[j] = ss[i];j++;
            }
            double period = 0; //利用distancematrix除以車速即可得到現在的時間
            double cost = 0.0;
            period += (startTOvertex[Sequence[0].StationID] / speed );//抵達第一站的時間 
            totalDistance += startTOvertex[Sequence[0].StationID]; //抵達第一站的距離
            int a = 0;Truck truck = new Truck(truckcapacity, truckcapacity/2);
            int[] temp = new int[totalStation];
            while(period <= timehorizon && a < totalStation - 1)
            {

                if (Sequence[a].Declinebymin < 0)
                {
                    temp[a] = truck.PickupDelivery(truckcapacity/2);
                    cost += HoldingShortageSurplus(period, Sequence[a].StationID,temp[a],ref truck);
                }
                else if (Sequence[a].Surplusbymin > 0)
                {
                    temp[a] = truck.PickupDelivery(-truckcapacity/2);
                    cost += HoldingShortageSurplus(period, Sequence[a].StationID,temp[a], ref truck);
                }
                
                period += distanceMatrix[Sequence[a].StationID, Sequence[a + 1].StationID] /speed;//到下一站的時間點
                totalDistance += distanceMatrix[Sequence[a].StationID, Sequence[a + 1].StationID];
                a++;
            }
            if(totalDistance >= speed * timehorizon) { totalDistance-= distanceMatrix[Sequence[a-1].StationID, Sequence[a].StationID]; }
            //因為上面最後一次部會做所以...
            if (Sequence[a].Declinebymin < 0 && period <= timehorizon)
            {
                temp[a] = truck.PickupDelivery(truckcapacity/2);
                cost += HoldingShortageSurplus(period, Sequence[a].StationID, temp[a], ref truck);
            }
            else if(Sequence[a].Surplusbymin > 0 && period <= timehorizon)
            {
                temp[a] = truck.PickupDelivery(-truckcapacity/2);
                cost += HoldingShortageSurplus(period, Sequence[a].StationID, temp[a], ref truck);
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
                    cost += HoldingShortageSurplus(timehorizon, Sequence[b].StationID, 0, ref truck);
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
            richTextBox.Text +=("\n" +"********performance********");
            richTextBox.Text +=  ("\n" + "total unsatisfied amount is :" + cost);
            richTextBox.Text += ("\n" + "total distance is :" + totalDistance );
            richTextBox.Text += ("\n" + "total pickup & delivery: " + totalpnd);
        }

        private void nearestInsert(object sender, EventArgs e)
        {
            totalDistance = 0;
            List<int> seq = new List<int>();
            double min = double.MaxValue; int last = -1;int next = -1;
            double[,] d = new double[totalStation, totalStation];
            //seq.Add(0); //must begin at 0
            for(int a = 0;a < totalStation; a++)
            {
                for(int b = 0;b < totalStation; b++)
                {
                    d[a, b] = distanceMatrix[a, b];
                }
            }
            for(int i = 1; i < totalStation; i++)
            {
                if (min > startTOvertex[i]) { min = startTOvertex[i]; last = i; }
            }
            //for(int j = 0;j < totalStation; j++)
            //{
            //    d[j, 0] = 0;
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
            } while (check(d,next));
            int[] routing = new int[seq.Count];
            seq.CopyTo(routing);
            double period = 0; //利用distancematrix除以車速即可得到現在的時間
            double cost = 0.0;
            Truck truck = new Truck(truckcapacity,truckcapacity/2);
            period += (startTOvertex[routing[0]] / speed );//抵達第一站的時間 
            totalDistance += startTOvertex[routing[0]]; //抵達第一站的距離         
            int c = 0;int[] temp = new int[totalStation];
            while (c < totalStation-1 && period <= timehorizon)
            {
                if (stations[routing[c]].Declinebymin < 0)
                {
                    temp[c] = truck.PickupDelivery(truckcapacity / 2);
                    cost += HoldingShortageSurplus(period, routing[c],temp[c], ref truck);
                }
                else
                {
                    temp[c] = truck.PickupDelivery(-truckcapacity / 2);
                    cost += HoldingShortageSurplus(period, routing[c], temp[c], ref truck);
                }
                period += distanceMatrix[routing[c], routing[c + 1]]/speed;//到下一站的時間點
                totalDistance += distanceMatrix[routing[c], routing[c + 1]];
                c++;
            }
            //因為上面最後一次部會做所以...
            if (stations[routing[c]].Declinebymin < 0 && period <= timehorizon)
            {
                temp[c] = truck.PickupDelivery(truckcapacity / 2);
                cost += HoldingShortageSurplus(period, routing[c], temp[c], ref truck);
            }
            else if(stations[routing[c]].Surplusbymin > 0 && period <= timehorizon)
            {
                temp[c] = truck.PickupDelivery(-truckcapacity / 2);
                cost += HoldingShortageSurplus(period, routing[c], temp[c], ref truck);
            }
            else { c--; }
            
            //全部跑完要回去depot
            if (c == totalStation - 1)
            {
                totalDistance += startTOvertex[routing[c]];
            }
            else
            {
                //沒有跑完的車站繼續算成本
                for (int b = c; b < totalStation - 1; b++)
                {
                    cost += HoldingShortageSurplus(timehorizon, routing[b], 0, ref truck);
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
            richTextBox.Text += "\n" + "********performance********";
            richTextBox.Text += ("\n" + "total unsatisfied amount is :" + cost);
            richTextBox.Text += ("\n" + "total distance is " + totalDistance);
            richTextBox.Text += ("\n" + "total pickup & delivery is " + tpnd);
           
        }
        private bool check(double[,] m,int n)
        {  bool flag = false;
            for (int i = 0; i < totalStation; i++)
            {
                if (m[n, i] != 0) { flag = true; break; }
            }
            return flag;
        }

        private void NBS(object sender, EventArgs e)
        {
            totalDistance = 0;
            Station[] ss = new Station[stations.Length];
            stations.CopyTo(ss,0);
            Array.Sort(ss);
            //找出不需要用送的站點作為補給站之用
            List<Station> supplyStation = new List<Station>();
            Priority_Queue.SimplePriorityQueue<Station> pq = new Priority_Queue.SimplePriorityQueue<Station>();
        
            for (int i = 0; i < ss.Length; i++)
            {
                if (ss[i].Rate != 0)
                {
                    pq.Enqueue(ss[i],1);
                }
                if (ss[i].Rate == 0)
                {
                    supplyStation.Add(ss[i]);
                }
            }
            
            int[] Sequence = new int[totalStation];
            for(int i = 0; i < Sequence.Length; i++)
            {
                Sequence[i] = -1;
            }
            //確定第一站為最需要先被服務的並且直接先決定卡車上的貨物數量
            //直接先決定到的時間點剛好為最需要被服務的剛好缺的時點
            int[] temp = new int[totalStation];
            double time = 0;
            double cost = 0;
            Truck truck = null;
            for(int k = ss.Length -1; k >-1 ;k--)
            {
                if (ss[k].Rate != 0)
                {
                    if (ss[k].Rate < 0) //rate < 0 => need to delivery
                    {
                        truck = new Truck(truckcapacity, truckcapacity);
                        temp[0] = truckcapacity / 2;
                        //cost = HoldingShortageSurplus(time, Sequence[0], truck.PickupDelivery(14));
                    }
                    else if (ss[k].Rate > 0)
                    {
                        truck = new Truck(truckcapacity, 0);
                        temp[0] = -truckcapacity / 2;                      
                        //cost = HoldingShortageSurplus(time, Sequence[0], truck.PickupDelivery(-14));
                    }
                    time = startTOvertex[ss[k].StationID] / speed;
                    Sequence[0] = ss[k].StationID;
                    temp[0] = truck.PickupDelivery(temp[0]);
                    cost += HoldingShortageSurplus(time, Sequence[0], temp[0], ref truck);
                    break;
                }
            }
           

            Station s = null;
            //開始建構解     
            int a = 1;bool flag = false;
            double min = double.MaxValue;double m = 0;int selected = -1;
            
            while (time <= timehorizon && a < totalStation)
            {
                while (pq.Count != 0)
                {
                    s = pq.Max();
                    pq.Remove(s); flag = false;
                    //s = pq.Dequeue();
                    if (Sequence.Contains(s.StationID) == false)
                    {
                        Sequence[a] = s.StationID;
                        break;
                    }
                    if (pq.Count == 0)
                    {
                        if(supplyStation.Count != 0)
                        {
                            s = supplyStation.Max();
                            supplyStation.Remove(s);
                        }
                        break;
                    }
                }
                min = double.MaxValue;
                //下一站是需要貨物的可是你卻沒有貨ㄌ
                if (s.Declinebymin < 0)
                {
                    //find truck lack of goods
                    if (truck.CurrentGoods == 0 && supplyStation.Count != 0)
                    {
                        min = double.MaxValue;
                        //find the nearest supply station
                        for (int i = 0; i < supplyStation.Count; i++)
                        {
                            m = distanceMatrix[Sequence[a - 1], supplyStation.ElementAt(i).StationID];
                            if (min > m)
                            {
                                min = m; selected = i;
                            }
                        }
                        temp[a] = truck.SupplyTruck(supplyStation.ElementAt(selected));
                        flag = true;
                        Sequence[a] = supplyStation.ElementAt(selected).StationID;
                        supplyStation.RemoveAt(selected);
                        if (a != stations.Length - 1)
                        {
                            Sequence[a + 1] = s.StationID;
                        }
                    }
                    if (flag == true)
                    {
                        time += distanceMatrix[Sequence[a - 1], Sequence[a]] / speed;//到下一站的時間點
                        totalDistance += distanceMatrix[Sequence[a - 1], Sequence[a]];
                        if (a != stations.Length - 1)
                        {
                            time += distanceMatrix[Sequence[a], Sequence[a + 1]] / speed;//到下一站的時間點
                            totalDistance += distanceMatrix[Sequence[a], Sequence[a + 1]];
                            temp[a + 1] = truck.PickupDelivery(truckcapacity / 2);
                            cost += HoldingShortageSurplus(time, Sequence[a + 1], temp[a + 1], ref truck);
                        }
                        a+=2;
                    }
                    else
                    {
                        temp[a] = truck.PickupDelivery(truckcapacity / 2);
                        cost += HoldingShortageSurplus(time, Sequence[a], temp[a], ref truck);
                    }
                    
                }    
                else
                {
                    //find truck full of goods
                    if (truck.CurrentSpace == 0 && supplyStation.Count != 0)
                    {
                        //find the nearest supply station
                        min = double.MaxValue;
                        for (int i = 0; i < supplyStation.Count; i++)
                        {
                            m = distanceMatrix[Sequence[a - 1], supplyStation.ElementAt(i).StationID];
                            if (min > m)
                            {
                                min = m; selected = i;
                            }

                        }             
                            temp[a] = truck.SupplyTruck(supplyStation.ElementAt(selected));
                            flag = true;
                            Sequence[a] = supplyStation.ElementAt(selected).StationID;
                            supplyStation.RemoveAt(selected);
                            if (a != stations.Length - 1)
                            {
                                Sequence[a + 1] = s.StationID;
                            }
                        
                    }
                    if (flag == true)
                    {
                        time += distanceMatrix[Sequence[a - 1], Sequence[a]] / speed;//到下一站的時間點
                        totalDistance += distanceMatrix[Sequence[a - 1], Sequence[a]];
                        if (a != stations.Length - 1)
                        {

                            time += distanceMatrix[Sequence[a], Sequence[a + 1]] / speed;//到下一站的時間點
                            totalDistance += distanceMatrix[Sequence[a], Sequence[a + 1]];
                            temp[a + 1] = truck.PickupDelivery(-truckcapacity / 2);
                            cost += HoldingShortageSurplus(time, Sequence[a + 1], temp[a + 1], ref truck);
                        }
                        a+=2;
                    }
                    else
                    {
                        temp[a] = truck.PickupDelivery(-truckcapacity / 2);
                        cost += HoldingShortageSurplus(time, Sequence[a], temp[a], ref truck);
                    }
                }
                if (flag == false)
                {
                    time += distanceMatrix[Sequence[a - 1], Sequence[a]]/speed;//到下一站的時間點
                    totalDistance += distanceMatrix[Sequence[a - 1], Sequence[a]];
                    a++;
                }        
            }
            if (a >= totalStation) { a = totalStation - 1; }
            //因為上面最後一次部會做所以...
            if (flag==false && stations[Sequence[a - 1]].Declinebymin < 0 && time <= timehorizon)
            {
                temp[a-1] = truck.PickupDelivery(truckcapacity / 2);
                cost += HoldingShortageSurplus(time, Sequence[a - 1], temp[a - 1], ref truck);
            }
            else if (flag == false && stations[Sequence[a - 1]].Surplusbymin > 0 && time <= timehorizon)
            {
                temp[a-1] = truck.PickupDelivery(-truckcapacity / 2);
                cost += HoldingShortageSurplus(time, Sequence[a - 1], temp[a - 1], ref truck);
            }
            else { a--; }

            //全部跑完要回去depot
            if (a == totalStation - 1)
            {
                totalDistance += startTOvertex[Sequence[a]];
            }
            else
            {
                //沒有跑完的車站繼續算成本
                for (int b = 0; b < pq.Count; b++)
                {
                    cost += HoldingShortageSurplus(timehorizon, pq.ElementAt(b).StationID, 0, ref truck);
                }
            }
            richTextBox.Text =  "routing path is depot ";
            for (int i = 0; i < a+1 ; i++)
            {
                richTextBox.Text += (Sequence[i] + " ");
            }

            double totalPandD = 0;
            richTextBox.Text += "\n" + "pickup & delivery amount: ";
            for (int i = 0; i < temp.Length; i++)
            {
                richTextBox.Text += temp[i].ToString() + "  ";
                totalPandD += Math.Abs(temp[i]);
            }
            richTextBox.Text += "\n" + "********performance********";
            richTextBox.Text += "\n" + "unsatisfied amount :" + cost;
            richTextBox.Text += ("\n" + "total distance is " + totalDistance);
            richTextBox.Text += "\n" + "Total pickup & delivery : " + totalPandD;
          
        }
    }
}
