using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BikeSharingSystem
{
    class HybridACOGA:GA<int>
    {
        protected int numberOfCity;
        protected int[][] solutions;      // Default solutions are of int type
        protected double dropPheromone = 0.01;       // Simplest method uses this amount to drop on each segment
        protected double dropMultiplier = 1.0; // Q value; the drop amount is    dropMultiplier/objective * dropPheromone
        protected int selectedIdx = 0;
        protected double[,] pheromone;             // Rectangular matrix (square for nxn) of the pheromone values
        protected double[,] heuristicValues;       // Matrix of the heuristic values, usually initialized from the problem
        protected ObjectiveFunctionDelegate getObjectiveValue = null;  // Function delegate for computing the objective value, provided by the user
        protected double[] probabilities;    // Variables to keep the compued probabilities of the candidate elements
        protected int[] candidateSet;  // The indices of the candidate elements
        protected int[] indicesOfAnts;        // Used in pheromone updating to hold indices of ants, when objective values are sorted
        protected double pheromoneFactor = 1.0;  // Power factor for pheromone value in computing probabilities of candidate elements
        protected double heuristicFactor = 1.0;  // Power factor for heuristic value in computing probabilities of candidate elements
        protected double evaporationRate = 0.1; // Pheromone evaporation rate
        double[,] FromToDistanceMatrix = null;
        double AverageDistance ;
        int[] soFarTheBestSoluiton;
        double[] ObjectiveValues;
        double ITB;
        double SFB = double.MaxValue;
        public int[][] Solution
        {
            get { return solutions; }
        }
        public int[] SoFarTheBestSoluiton
        {
            get { return soFarTheBestSoluiton; }
        }
        public void reset(int P)
        {
            indicesOfAnts = new int[P];
            solutions = new int[P][];
            for (int i = 0; i < P; i++)
            {
                indicesOfAnts[i] = i;
                solutions[i] = new int[numberOfCity];
                for (int j = 0; j < numberOfCity; j++)
                {
                    solutions[i][j] = j;
                }
                KnuthShuffle(solutions[i]);
            }
        }
        public HybridACOGA(int numberOfVariables, OptimizationType opType, GA<int>.ObjectiveFunctionDelegate objectFunction, double[,] distance) : base(numberOfVariables, opType, objectFunction)
        {
            getObjectiveValue = objectFunction;
            numberOfCity = numberOfVariables;
            FromToDistanceMatrix = distance;
            // Allocate memory for pheromone matrix and heuristic value matrix
            // Heuristic values in the matrix are computed and assigned via 
            // the specified function delegate
            pheromone = new double[numberOfCity, numberOfCity];
            indicesOfAnts = new int[populationSize];
            solutions = new int[populationSize][];
            soFarTheBestSoluiton = new int[numberOfCity];
            for (int r = 0; r < numberOfCity; r++)
            {
                for (int c = 0; c < numberOfCity; c++)
                {
                    pheromone[r, c] = 0.01;
                }
            }
            for (int r = 0; r < numberOfCity; r++)
            {
                for (int c = 0; c < numberOfCity; c++)
                {
                    AverageDistance += FromToDistanceMatrix[r, c];
                }
            }
            AverageDistance = AverageDistance / (numberOfCity * numberOfCity);

            heuristicValues = new double[numberOfCity, numberOfCity];
            for (int r = 0; r < numberOfCity; r++)
            {
                for (int c = 0; c < numberOfCity; c++)
                {
                    heuristicValues[r, c] = DistanceInverseHeuristic(r, c);
                }
            }
            for (int i = 0; i < populationSize; i++)
            {
                indicesOfAnts[i] = i;
                solutions[i] = new int[numberOfCity];
                for(int j = 0; j < numberOfCity; j++)
                {
                    solutions[i][j] = j;
                }
                KnuthShuffle(solutions[i]);
            }            
         
            // Allocate memory for the arries used in ACO, whose length is 
            // depend on the number of variables. In contrast, those arraies 
            // with length depending on the number of ants are allocated in 
            // reset function.
            probabilities = new double[numberOfCity];
            candidateSet = new int[numberOfCity];
            ObjectiveValues = new double[populationSize];
            // Set drop amount multiplier; which is the avarage length
            //dropMultiplier = dropPheromone * dropMultiplier / getObjectiveValue(objectiveValues);

        }

        public void KnuthShuffle(int[] array)
        {
            int temp;
            for (int i = 0; i < (array.Length - 2) / 2; i++)
            {
                int j = rnd.Next(i, (array.Length - 2) / 2); // Don't select from the entire array on subsequent loops
                temp = array[i]; array[i] = array[j]; array[j] = temp;
            }
        }

        public double DistanceInverseHeuristic(int i, int j)
        {
            double B = AverageDistance / 5.0;

            return FromToDistanceMatrix[i, j] < 0.000001 ? 0.0 : B /FromToDistanceMatrix[i, j];
        }

        public virtual void antsConstructSolutions()
        {
            // Set the selection mode
            int elementLeft = -1;
            double[] cumulative = new double[numberOfCity];

            

            // Loop through each ant to construt each solution

            for (int i = 0; i < populationSize; i++)
            {

                // Prepare candidate set, using indicesOfVariables array.
                for (int j = 0; j < candidateSet.Length; j++)
                {

                    candidateSet[j] = j;  //準備城市候選
                }
                elementLeft = numberOfCity; //城市ㄉ指標

                selectedIdx = rnd.Next(elementLeft);

                solutions[i][0] = candidateSet[selectedIdx]; //指定起始城市點
                elementLeft--; // The candidate reduced by one
                candidateSet[selectedIdx] = candidateSet[elementLeft];// Discard the selected index by moving back the last candidate to this position and then
                for (int q = 1; q < numberOfCity; q++)//開始找出下一個城市和前一個城市的相配機率
                { // Compute the probabilities for each candidate to succeed the previous element

                    for (int b = 0; b < elementLeft; b++)
                    {
                        //求算和下一個城市ㄉ適合機率
                        probabilities[b] = Math.Pow(pheromone[solutions[i][q - 1], candidateSet[b]], pheromoneFactor) * Math.Pow(heuristicValues[solutions[i][q - 1], candidateSet[b]], heuristicFactor);

                    }

                        // Repeatly select the next element to succeed the previous one, until the solution is completed.
                        // Deterministic method, find the best candidate, by evaluating their selection probability
                        //// Directly select the element with the largest probability 
                        //Array.Sort(probabilities, candidateSet,0,elementLeft);
                        ////Array.Reverse(probabilities); Array.Reverse(candidateSet);
                        double m = double.MinValue;
                        selectedIdx = -1;
                        //找到機率大的城市做為下一個候選城市
                        for (int w = 0; w < elementLeft; w++)
                        {
                            if (probabilities[w] > m)
                            {
                                selectedIdx = w;
                                m = probabilities[w];
                            }

                        }             
                    //已挑到下一個城市
                    solutions[i][q] = candidateSet[selectedIdx]; //將挑到的下一個城市指定給下一個位子
                    elementLeft--;//指標向左移一個
                    candidateSet[selectedIdx] = candidateSet[elementLeft];//將之前指標指向的那個城市index替換被選到ㄉ
                }
            }

        }

        public virtual void ComputeObjectiveValues()
        {
            // Compute the objective value for each solution contructed by the ants
            for (int i = 0; i < populationSize; i++)
            {
                ObjectiveValues[i] = getObjectiveValue(solutions[i]);  //表示拿到第i隻螞蟻ㄉsolution
            }
            // Sort the objective values and ant indices in orders of superior to inferior
            Array.Sort(ObjectiveValues, solutions, 0, populationSize);
            ITB = ObjectiveValues[0];
            soFarTheBestSoluiton = solutions[0];
            if (ITB < SFB)
            {
                SFB = ITB;
            }
        }

        public virtual void updatePheromone()
        {
            int k;
            // Loop throught each solution, identify the element indices for each segment
            // Add pheromone drop amount to the identified segment
            k = soFarTheBestSoluiton[numberOfCity - 1];
            for (int j = 0; j < numberOfCity; j++)
            {
                pheromone[k, soFarTheBestSoluiton[j]] += dropPheromone * dropMultiplier / SFB;
                k = soFarTheBestSoluiton[j];
            }
            //for (int i = 0; i < populationSize; i++)
            //{
            //    k = solutions[i][numberOfCity - 1];
            //    for (int j = 0; j < numberOfCity; j++)
            //    {
            //        pheromone[k, solutions[i][j]] += dropPheromone * dropMultiplier / ObjectiveValues[i];
            //        k = solutions[i][j];
            //    }
            //}

            // Evaporate the pheromone for all pheromone values
            for (int r = 0; r < pheromone.GetLength(0); r++)
                for (int c = 0; c < pheromone.GetLength(1); c++)
                    pheromone[r, c] *= (1.0 - evaporationRate);
        }
    }
}
