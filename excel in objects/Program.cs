using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace excel_in_objects
{
    public class Cell
    {
        public virtual string Write()
        {
            return "";
        }

        public static void Evaluate(int row, int column)
        {
            Stack stackOfCells = new Stack();
            Stack stackOfNumbers = new Stack();

            Formula cell = (Formula)Program.excelTable[row][column];

            cell.visited = true;

            stackOfCells.Push((row,column));
            stackOfCells.Push(cell.operation);
            stackOfCells.Push(cell.right);
            stackOfCells.Push(cell.left);


            while (stackOfCells.Count > 0)
            {
                var popped = stackOfCells.Pop();
                if (!(popped is char))
                {
                    (int, int) poppedTyped = ((int, int))popped;
                    if (poppedTyped.Item1 > Program.excelTable.Count - 1 || poppedTyped.Item2 > Program.excelTable[poppedTyped.Item1].Length - 1)
                        stackOfNumbers.Push(0);
                    else if (Program.excelTable[poppedTyped.Item1][poppedTyped.Item2] is Formula)
                    {
                        if (((Formula)Program.excelTable[poppedTyped.Item1][poppedTyped.Item2]).visited)
                        {
                            stackOfNumbers.Push(true); //true znamená ciklus
                        }
                        else
                        {
                            Formula newCell = (Formula)Program.excelTable[poppedTyped.Item1][poppedTyped.Item2];
                            stackOfCells.Push(popped);
                            stackOfCells.Push(newCell.operation);
                            stackOfCells.Push(newCell.right);
                            stackOfCells.Push(newCell.left);
                        }

                    }
                    else if (Program.excelTable[poppedTyped.Item1][poppedTyped.Item2] is Number)
                    {
                        Number newNumber = (Number)Program.excelTable[poppedTyped.Item1][poppedTyped.Item2];
                        stackOfNumbers.Push(newNumber.number);
                    }

                    else if (Program.excelTable[poppedTyped.Item1][poppedTyped.Item2] is Brackets)
                        stackOfNumbers.Push(0);

                    else
                        stackOfNumbers.Push(false); //false znamená akýkoľvek error
                }
                else
                {
                    char poppedTyped = (char)popped;
                    var firstValue = stackOfNumbers.Pop();
                    var secondValue = stackOfNumbers.Pop();
                    (int, int) address = ((int, int))stackOfCells.Pop();

                    if (firstValue is bool && secondValue is bool)
                    {
                        bool firstValueBool = (bool)firstValue;
                        bool secondValueBool = (bool)secondValue;
                        if (firstValueBool || secondValueBool)
                        {
                            stackOfNumbers.Push(true);
                            Program.excelTable[address.Item1][address.Item2] = new Error(3);
                        }
                        else
                        {
                            stackOfNumbers.Push(false);
                            Program.excelTable[address.Item1][address.Item2] = new Error(1);
                        }
                    }
                    else if (firstValue is bool)
                    {
                        bool firstValueBool = (bool)firstValue;
                        if (firstValueBool)
                        {
                            stackOfNumbers.Push(true);
                            Program.excelTable[address.Item1][address.Item2] = new Error(3);
                        }
                        else
                        {
                            stackOfNumbers.Push(false);
                            Program.excelTable[address.Item1][address.Item2] = new Error(1);
                        }
                    }
                    else if (secondValue is bool)
                    {
                        bool secondValueBool = (bool)secondValue;
                        if (secondValueBool)
                        {
                            stackOfNumbers.Push(true);
                            Program.excelTable[address.Item1][address.Item2] = new Error(3);
                        }
                        else
                        {
                            stackOfNumbers.Push(false);
                            Program.excelTable[address.Item1][address.Item2] = new Error(1);
                        }
                    }

                    else
                    {

                        int number;
                        switch (poppedTyped)
                        {

                            case '+':
                                number = (int)firstValue + (int)secondValue;
                                stackOfNumbers.Push(number);
                                Program.excelTable[address.Item1][address.Item2] = new Number(number);
                                break;
                            case '-':
                                number = (int)secondValue - (int)firstValue;
                                stackOfNumbers.Push(number);
                                Program.excelTable[address.Item1][address.Item2] = new Number(number);
                                break;
                            case '*':
                                number = (int)firstValue * (int)secondValue; ;
                                stackOfNumbers.Push(number);
                                Program.excelTable[address.Item1][address.Item2] = new Number(number);
                                break;
                            case '/':
                                if ((int)firstValue == 0)
                                {
                                    Program.excelTable[address.Item1][address.Item2] = new Error(2);
                                }
                                else
                                {
                                    number = (int)secondValue / (int)firstValue;
                                    stackOfNumbers.Push(number);
                                    Program.excelTable[address.Item1][address.Item2] = new Number(number);
                                }
                                break;
                        }
                    }
                }

            }

        }
        
        public static Cell CreateCell(string cellString)
        {
            if (cellString[0] == '=')
            {
                char operation = ' ';
                bool hasOperator = false;
                bool hasMoreThanOneOperator = false;
                for (int i = 0; i < cellString.Length; i++)
                {
                    if (cellString[i] == '-' || cellString[i] == '+' || cellString[i] == '*' || cellString[i] == '/')
                    {

                        if (hasOperator)
                        {
                            hasMoreThanOneOperator = true;
                            break;
                        }
                        else
                        {
                            hasOperator = true;
                            operation = cellString[i];
                        }
                    
                    }
                }

                if (hasOperator)
                {
                    bool hasUpper = false;
                    if (!hasMoreThanOneOperator)
                    {
                        string[] leftAndRightCell = cellString.Split(operation);
                        for (int i = 1; i < leftAndRightCell[0].Length; i++)
                        {
                            if (char.IsUpper(leftAndRightCell[0][i]))
                                hasUpper = true;
                            else if (int.TryParse(leftAndRightCell[0].Substring(i), out int leftNumber) && hasUpper)
                            {
                                hasUpper = false;

                                for (int j = 0; j < leftAndRightCell[1].Length; j++)
                                {
                                    if (char.IsUpper(leftAndRightCell[1][j]))
                                        hasUpper = true;
                                    else if (int.TryParse(leftAndRightCell[1].Substring(j), out int rightNumber) && hasUpper)
                                        return new Formula(leftAndRightCell[0].Substring(1,i-1), leftNumber, leftAndRightCell[1].Substring(0,j), rightNumber, operation);
                                    else 
                                        return new Error(5);
                                }
                            }
                            else 
                                return new Error(5);
                        }
                    }
                    return new Error(5);
                }
                else
                    return new Error(4);
            }

            else if (int.TryParse(cellString, out int number))
                return new Number(number);
            else if (cellString == "[]")
                return new Brackets();
            else
                return new Error(0);
            
        }
    }

    public class Number : Cell
    {
        public int number;

        public Number(int number)
        {
            this.number = number;
        }

        public override string Write()
        {
            return number.ToString();
        }
    }

    public class Error : Cell
    {
        public int errorCode;
        //0 - #INVVAL
        //1 - #ERROR 
        //2 - #DIV0
        //3 - #CYCLE
        //4 - #MISSOP
        //5 - #FORMULA

        public Error(int errorCodeNumer)
        {
            errorCode = errorCodeNumer;
        }

        public override string Write()
        {
            switch (errorCode)
            {
                case 0:
                    return "#INVVAL";
                case 1:
                    return "#ERROR";
                case 2:
                    return "#DIV0";
                case 3:
                    return "#CYCLE";
                case 4:
                    return "#MISSOP";
                default:
                    return "#FORMULA";
            }
        }
    }

    public class Formula : Cell
    {
        public char operation;
        public (int, int) left;
        public (int, int) right;
        public bool visited;

        public Formula(string leftLetters, int leftNumber, string rightLetters, int rightNumber, char operation)
        {
            visited = false;
            int leftColumn = -1;
            int rightColumn = -1;

            for (int i = 0; i < leftLetters.Length; i++)
                leftColumn += (((int)leftLetters[leftLetters.Length - 1 - i]) - 64) * (int)Math.Pow(26, (i));
            

            for (int i = 0; i < rightLetters.Length; i++)
                rightColumn += (((int)rightLetters[rightLetters.Length - 1 - i]) - 64) * (int)Math.Pow(26, (i));

            left = (leftNumber - 1,leftColumn);
            right = (rightNumber - 1,rightColumn);

            this.operation = operation;

        }
    }

    public class Brackets : Cell
    {
        public Brackets()
        {
        }

        public override string Write()
        {
            return "[]";
        }
    }
    
    class Program
    {
        
        public static List<Cell[]> excelTable = new List<Cell[]>();

        public static void Write(string file)
        {
            using (StreamWriter sw = new StreamWriter(file))
            {
                for (int i = 0; i < excelTable.Count - 1; i++)
                {
                    for (int j = 0; j < excelTable[i].Length - 1; j++)
                    {
                        sw.Write(excelTable[i][j].Write());
                        sw.Write(' ');
                    }
                    if (excelTable[i].Length > 0)
                        sw.WriteLine(excelTable[i][excelTable[i].Length - 1].Write());
                    else
                        sw.WriteLine();
                }

                for (int j = 0; j < excelTable[excelTable.Count - 1].Length - 1; j++)
                {
                    sw.Write(excelTable[excelTable.Count - 1][j].Write());
                    sw.Write(' ');
                }
                try { sw.Write(excelTable[excelTable.Count - 1][excelTable[excelTable.Count - 1].Length - 1].Write()); }
                catch { } 
                
            }
        }
        static void Main(string[] args)
        {
            /**/
            if (args.Length != 2)
                Console.WriteLine("Argument Error");
            else
            {
                try
                {
                    using (StreamReader sr = new StreamReader(args[0]))
                    {
                        string line;
                        int row = 0;
                        int column = 0;
                        List<(int, int)> addressesToProcess = new List<(int, int)>();

                        while ((line = sr.ReadLine()) != null)
                        {
                            string[] strings = line.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            Cell[] cells = new Cell[strings.Length];

                            for (int i = 0; i < cells.Length; i++)
                            {
                                Cell cell = Cell.CreateCell(strings[i]);
                                cells[i] = cell;
                                if (cell is Formula)
                                    addressesToProcess.Add((row, column));
                                ++column;

                            }
                            column = 0;
                            ++row;
                            excelTable.Add(cells);
                        }

                        for (int i = 0; i < addressesToProcess.Count; i++)
                        {
                            if (excelTable[addressesToProcess[i].Item1][addressesToProcess[i].Item2] is Formula)
                                Cell.Evaluate(addressesToProcess[i].Item1, addressesToProcess[i].Item2);
                        }

                        Write(args[1]);
                    }
                }
                catch
                {
                    Console.WriteLine("File Error");
                }
            }
        }
    }
}
