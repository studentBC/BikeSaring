using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BikeSharingSystem
{
    public class GA<T>
    {
        // A nested delegate (data Type) declaration defining a function returning a double value 
        // while taking an argument of an array of type T
        public delegate double ObjectiveFunctionDelegate(T[]chromosome);

        // Problem associated data, which are specified via constructor
        // Delegate (function pointer in C) for the user-supplied objective evalution function.
        public ObjectiveFunctionDelegate GetObjectiveValueFunction;
        protected OptimizationType optimizationType = OptimizationType.Min;
        protected int numberOfGenes; // The length of a chromosome 1
        protected int anotherGenes; // the length of a chromosome (卡車出發時間點+初始貨物數量)
        protected double goalValue;
        protected int BestOne;
        protected int populationSize = 20;
        protected int numberOfCrossoveredChildren;
        protected int numberOfMutatedChildren;
        protected double crossoverRate = 0.8;
        protected double mutationRate = 0.3;
        protected int iterationLimit = 1000; // The default stopping condition
        protected int iterationCount;        // Current iteration index
        SelectionMode selectionMode = SelectionMode.Deterministic;
        PermutationCrossover crossoverType = PermutationCrossover.NpointCut;

        protected T[][] chromosomes;         // Parameterized type, specified by the derived class
        protected double[] objectiveValues;  // Positive or negative real numbers
        protected double[] fitnessValues;    // Must be positive values

        protected double iterationBestObjective;   // The best value obtained in the current iteration
        protected double iterationAverage;         // The objective average obtained in the current iteration

        protected double soFarTheBestObjective; //in fact it is private
        protected T[] soFarTheBestSolution;

        // Used to stochastically select pairs of parents for crossover operation.
        // Used in deterministic selection operation to sort the chromosome indices for ordered fitness
        // Used in stochastic selection operation to store the indices of chosen chromosomes
        protected int[] indices;     // Store indices of chromosome, which are subject to shuffle, sorting, inversing operations

        protected Random rnd = new Random();  // Shared by all derived classes


        #region PropertiesRegion
        [Category("GA Setting"), Description("染色體數量，建議值是不低於變數數量的1/4。")]
        public int PopulationSize
        {
            get
            {
                return populationSize;
            }
            set
            {
                if (populationSize < 2) return;
                populationSize = value;
            }
        }

        // ...

        [Browsable(false)]
        public double SoFarTheBestObjective
        {
            get
            {
                return soFarTheBestObjective;
            }
        }

        [Browsable(false)]
        public T[] SoFarTheBestSolution
        {
            set
            {
                soFarTheBestSolution = value;
            }
            get
            {
                return soFarTheBestSolution;
            }
        }
        [Browsable(false)]
        public int IterationCount
        {
            get
            {
                return iterationCount;
            }
        }
        [Category("GA Setting"), Description("the more the better ")]
        public int IterationLimit
        {
            get
            {
                return iterationLimit;
            }

            set
            {
                if (value > 1)
                { iterationLimit = value; }
            }
        }
        //[Category("GA Setting"), Description("交配的方式")]
        //public PermutationCrossover CrossOverMode
        //{
        //    get
        //    {
        //        return crossoverType;
        //    }
        //    set
        //    {
        //        crossoverType = value;
        //    }
        //}
        [Category("GA Setting"), Description("篩選的模式")]
        public SelectionMode SelectionMode
        {
            get
            {
                return selectionMode;
            }
            set
            {
                selectionMode = value;
            }
        }
        [Category("GA Setting"),Description("突變機率建議不大於0.5")]
        public double MutationRate
        {
            set
            {
                if (mutationRate <= 0.5)
                { mutationRate = value; }
                else { mutationRate = 0.5; }
            }
            get { return mutationRate; }
        }
        [Browsable(false)]
        public double IterationBestObjective
        {
            get
            {
                return iterationBestObjective;
            }
        }

        [Browsable(false)]
        public double IterationAverage
        {
            get
            {
                return iterationAverage;
            }
        }
        #endregion


        /// <summary>
        ///  To employ a GA solver, user must provide the number of variables, the optimization type, and a function delegate
        ///  that compute and return the objective value for a given solution.
        /// </summary>
        /// <param name="numberOfVariables"> Number of variables of the problem</param>
        /// <param name="opType"> The optimization problem type </param>
        /// <param name="objectFunction"> The function delegate that computer the objective value for a given solution </param>
        public GA(int numberOfVariables, OptimizationType opType, GA<T>.ObjectiveFunctionDelegate objectFunction)
        {
            numberOfGenes = numberOfVariables;
            optimizationType = opType;
            if (objectFunction == null) throw new Exception("You must prepare an objective function.");
            GetObjectiveValueFunction = objectFunction;
        }


        /// <summary>
        ///  This function reallocate memeory for the GA computation subject to newly
        ///  specified properties; e.g, population size. In addition, the initial population
        ///  of chromosomes are initialized.
        /// </summary>
        T[,] temp;
        public virtual void reset()
        {
            // Allocate memory for gene related data
            chromosomes = new T[populationSize * 3][];
            indices = new int[chromosomes.Length];
            temp = new T[populationSize, numberOfGenes];
            // Initialize the initial population
            for (int i = 0; i < chromosomes.Length; i++)
            {
                chromosomes[i] = new T[numberOfGenes];
            }

            // Reset computation realted variables
            soFarTheBestSolution = new T[numberOfGenes];
            objectiveValues = new double[chromosomes.Length];
            fitnessValues = new double[chromosomes.Length];            
            initializePopulation();
            for (int j = 0; j < populationSize; j++)
            {
                objectiveValues[j] = GetObjectiveValueFunction(chromosomes[j]);
            }
            iterationCount = 0;
            soFarTheBestObjective = ((optimizationType == OptimizationType.Min) ? double.MaxValue : double.MinValue);

        }


        /// <summary>
        ///  This function setup the indices from 0 to upLimit-1 in indice array (int[] indices).
        ///  Then, shuffle their orders randomly. 
        ///  This function is called to shuffle the indice orders of parent population to support 
        ///  pair-wise crossover operation. If x pairs of parents are to be crossovered, then
        ///  the first 2x indices are the chromosome indices of the x pair parents.
        /// </summary>
        /// <param name="upLimit"></param>
        /// 

        protected void randomizeIndices(int upLimit)
        {
            for (int i = 0; i < populationSize*3; i++) { indices[i] = i; }
            for (int j = upLimit-1; j > 0; j--)   //做swap的洗牌
            {
                int temp = indices[j];
                int pos = rnd.Next(j + 1);
                indices[j] = indices[pos];
                indices[pos] = temp; 
            }
        }

        /// <summary>
        ///  Called in reset function. Overriden by the derived classes to fill-in
        ///  populationSize chromosomes with gene values of their data types.
        /// </summary>
        public virtual void initializePopulation()
        {

        }


        /// <summary>
        ///  Default method that carryout the whole GA computation without any interruption.
        /// </summary>
        public virtual void executeToEnd()
        {
            do
            {
                executeOneIteration();
            } while (!terminationConditionMet());
        }

        /// <summary>
        ///  A function that determine wether stopping condition is met. By default, the iteration 
        ///  limit is used and checked for termination. Derived calles can override it.
        /// </summary>
        /// <returns></returns>
        public virtual bool terminationConditionMet()
        {
            if (iterationCount < iterationLimit) return false;
            else return true;
        }


        /// <summary>
        ///  Standard GA computation procedure. However, derived classes may override it.
        /// </summary>
        public virtual void executeOneIteration()
        {
            // Crossover operation
            performCrossoverOperation();
            // Mutation operation
            performMutateOperation();
            // Evaluate all objectives 
            computeObjectiveValues();
            // Transform objectives to fitness values
            setFitnessFromObjectives();
            // Selection
            performSelectionOperation();

            iterationCount++;
        }


        /// <summary>
        ///  Standard function that evaluates original objective values for parent and children chromosomes.
        ///  During the computation, iteration best is identified and checked with the so far the best.
        ///  The so far the best objective and solution will be updated, if the iteration best surpass its value.
        ///  Specifically, this function calls the user-supplied objective value evalution function delegate to
        ///  evaluate each chromosome and put value to objectiveValues array. 
        /// </summary>
        public virtual void computeObjectiveValues()
        {
            int parentPlusChildren = populationSize + numberOfCrossoveredChildren + numberOfMutatedChildren;
          //  parentPlusChildren = chromosomes.Length ;
            for (int i = 0; i < parentPlusChildren ; i++)
            {
                objectiveValues[i] = GetObjectiveValueFunction(chromosomes[i]);
            }


            switch (optimizationType)
            {
                case OptimizationType.Min:
                    {
                        iterationBestObjective = double.MaxValue;
                        iterationAverage = 0.0;
                        for (int j = 0; j < parentPlusChildren; j++)
                        {
                            iterationAverage += objectiveValues[j];

                            if (objectiveValues[j] < iterationBestObjective)
                            {
                                iterationBestObjective = objectiveValues[j];
                                BestOne = j;
                            }
                        }
                        iterationAverage /= (double)parentPlusChildren;

                        if (iterationBestObjective < soFarTheBestObjective)
                        {
                            soFarTheBestObjective = iterationBestObjective;
                            for (int k = 0; k < numberOfGenes; k++)
                            {
                                soFarTheBestSolution[k] = chromosomes[BestOne][k];
                            }
                        }
                        break;
                    }
                case OptimizationType.Max:
                    {
                        iterationBestObjective = double.MinValue;
                        iterationAverage = 0.0;
                        for (int i = 0; i < parentPlusChildren; i++)
                        {
                            iterationAverage += objectiveValues[i];

                            if (objectiveValues[i] > iterationBestObjective)
                            {
                                iterationBestObjective = objectiveValues[i];
                                BestOne = i;
                            }

                        }
                        iterationAverage /= (double)parentPlusChildren;

                        if (iterationBestObjective > soFarTheBestObjective)
                        {
                            this.soFarTheBestObjective = this.iterationBestObjective;
                          
                                for (int i = 0; i < numberOfGenes; i++)
                                {
                                    soFarTheBestSolution[i] = chromosomes[BestOne][i];
                                }
                        }
                        break;
                    }

            }
        }

        /// <summary>
        ///  This function convert original objective values into positive fitness values, such that
        ///  the better chromosome receives the larger amount of fitness. Notice that the worest one
        ///  still receive the least amount of positive fitness value.
        ///  Sepcifically, the function transform each value in objectiveValues array to the value in
        ///  fitnessValues array.
        /// </summary>
        public void setFitnessFromObjectives()
        {
            int total = populationSize + numberOfCrossoveredChildren + numberOfMutatedChildren;
            double max = double.MinValue, min = double.MaxValue;int[] pd = new int[objectiveValues.Length];
            for (int i = 0; i < total; i++)
            {
                for (int a = (numberOfGenes - 2) / 2; a < numberOfGenes - 1; a++)
                {
                    pd[i] += Math.Abs(Convert.ToInt32(chromosomes[i][a]));
                    pd[i] = pd[i]/2;
                }
                if (objectiveValues[i]+pd[i] < min)
                {
                    min = objectiveValues[i]+pd[i];
                }

                if (objectiveValues[i] + pd[i] > max)
                {
                    max = objectiveValues[i] + pd[i];
                }       
            }
            double bound;
            switch (optimizationType)
            {
                case OptimizationType.Min:
                    for (int i = 0; i < total; i++)
                    {
                      
                        bound = max + (max - min) * 0.01;
                        fitnessValues[i] = (bound - (objectiveValues[i] + pd[i]) / (bound - min));
                        indices[i] = i;
                    }
                    break;
                case OptimizationType.Max:
                    for (int i = 0; i < total; i++)
                    {
                        bound = min - (max - min) * 0.01;
                        fitnessValues[i] = (bound - (objectiveValues[i] + pd[i]) / (bound - max));
                        indices[i] = i;
                    }
                    break;
            }

        }

        /// <summary>
        ///  Standard crossover operation in a GA iteration. With the help of a shuffled index array (indices array)
        ///  parent chromosomes are paired for crossover operation.
        ///  This standard function calls derived class overriden generateAPairOfCrossoveredOffspring() function to 
        ///  let that function access parent chromosome (via indices) and set gene values for the children chromosome
        ///  (via indices).
        /// </summary>
        public virtual void performCrossoverOperation()
        {
            // Calculate the number of crossovered chromosomes (must be even number)
            numberOfCrossoveredChildren = (int)(crossoverRate * populationSize);
            if (numberOfCrossoveredChildren % 2 == 1) numberOfCrossoveredChildren++;
            if (numberOfCrossoveredChildren > populationSize) numberOfCrossoveredChildren = populationSize;
            // Randomly sort the parent indices to select the parents for pair-wise crossover
            randomizeIndices(populationSize);
            // For a pair of parent indices, prepare the successive children indices, 
            // call  generateAPairOfCrossoveredOffspring() to produce crossovered offspring
            int child1 = populationSize, child2 = populationSize + 1;
            for (int i = 0; i < numberOfCrossoveredChildren; i += 2)
            {
                generateAPairOfCrossoveredOffspring(indices[i], indices[i + 1], child1, child2);
                child1 += 2;
                child2 += 2;
            }   //避免選到相同的父母
        }


        /// <summary>
        ///  Given two parent indices, and two children indices, this function perform crossover operation. The gene values 
        ///  of the children will be set by this function. This function must be overriden by the derived classes.
        /// </summary>
        /// <param name="fartherIdx"> index of farther chromosome </param>
        /// <param name="motherIdx"> index of mother chromosome  </param>
        /// <param name="child1Idx"> index of child 1 chromosome </param>
        /// <param name="child2Idx"> index of child 2 chromosome </param>
        public virtual void generateAPairOfCrossoveredOffspring(int fartherIdx, int motherIdx, int child1Idx, int child2Idx)
        {
        }

        /// <summary>
        ///  This function conducts one of the primary operaion in GA compution. Since different GA codings had different
        ///  mutation operations, no standard mutation operation is available.
        ///  Derived class must override this function.
        /// </summary>
        public virtual void performMutateOperation()
        {
        }


        /// <summary>
        ///  This function provide standard GA selection operation. However, it allowed derived classes to override it.
        ///  Two selection modes are provided in this funciton: deterministic and stochastic.
        /// </summary>

        public virtual void performSelectionOperation()
        {
            int Totalchromosons = populationSize + numberOfCrossoveredChildren + numberOfMutatedChildren;
            if (selectionMode == SelectionMode.Deterministic)
            {
              //  this.randomizeIndices(Totalchromosons);//why ??
                Array.Sort<double, int>(fitnessValues, indices, 0, Totalchromosons);
                Array.Reverse(indices, 0, Totalchromosons);
               // Array.Reverse(fitnessValues,0,Totalchromosons);
            }
            else if (selectionMode == SelectionMode.Stochastic)
            {              
                double wheelarea = 0.0;double hit = 0;
                for (int l = 0; l < Totalchromosons; l++)
                {
                    fitnessValues[l]+=wheelarea;
                    wheelarea = fitnessValues[l];
                }

                for (int i = 0; i < populationSize; i++)
                {
                    hit = rnd.NextDouble() * wheelarea;
                    for (int j = 0; j < Totalchromosons - 1; j++)
                    {
                        if (hit<= fitnessValues[1])
                            indices[i] = 0;
                        else if (fitnessValues[j] < hit && fitnessValues[j + 1] >= hit)
                        {
                            indices[i] = j + 1;
                        }
                    }
                }
            }
            Array.Sort(indices, 0, populationSize);
            //amend and provided by professor 難道可以信?!!!
            if (selectionMode == SelectionMode.Deterministic)
            {
                for (int i = 0; i < populationSize; i++)
                {
                    objectiveValues[i] = objectiveValues[indices[i]];
                    for (int j = 0; j < numberOfGenes; j++)
                    {
                        chromosomes[i][j] = chromosomes[indices[i]][j];
                    }
                }
            }
            else
            {
                for (int i = 0; i < populationSize; i++)
                {
                    bool reserved = false;
                    for (int j = 0; j < populationSize; j++)
                    {
                        if (i == indices[j])
                        {
                            indices[j] = -1;
                            reserved = true;
                            break;
                        }
                    }
                    if (!reserved)
                    {
                        for (int k = 0; k < populationSize; k++)
                        {
                            if (indices[k] < 0) continue;
                            for (int m = 0; m < numberOfGenes; m++)
                            {
                                chromosomes[i][m] = chromosomes[indices[k]][m];
                            }
                            objectiveValues[i] = objectiveValues[indices[k]];
                            indices[k] = -1;
                            break;
                        }
                    }
                }
            }
            //for (int i = 0; i < populationSize; i++)
            //{
            //    bool reserved = false;
            //    for (int j = 0; j < populationSize; j++)
            //    {
            //        if (i == indices[j])
            //        {
            //            indices[j] = -1;
            //            reserved = true;
            //            break;
            //        }
            //    }
            //    if (!reserved)
            //    {
            //        int chosen = -1;
            //        for (int k = 0; k < populationSize; k++)
            //        {
            //            if (!(indices[k] < 0))
            //            {
            //                chosen = indices[k];
            //                indices[k] = -1;
            //                break;
            //            }
            //        }
            //        for (int m = 0; m < numberOfGenes; m++)
            //        {
            //            chromosomes[i][m] = chromosomes[chosen][m];
            //        }
            //        objectiveValues[i] = objectiveValues[chosen];
            //    }
            //}

        }
    }


    /// <summary>
    ///  Type of optimization problem.
    /// </summary>
    public enum OptimizationType { Min, Max}

    /// <summary>
    ///  Type of GA selection procedure
    /// </summary>
    public enum SelectionMode { Deterministic, Stochastic }
    public enum PermutationMutation { RandomAssignment , ReciprocalExchange,Inversion, TSPHeuristic ,Heuristic}
}
