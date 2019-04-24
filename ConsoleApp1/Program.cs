using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp1
{
	class Program
	{
		static void Main(string[] args)
		{
			Excel.Application xlApp = new Excel.Application();
			Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\JakeA\source\repos\ConsoleApp1\Trivia-Printable.xlsx");
			Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
			Excel.Range xlRange = xlWorksheet.UsedRange;

			int rowCount = 1629;
			int colCount = 7;
			int i = 1;
			int userScore = 0;

			Random rand = new Random();


			Console.WriteLine("Who's ready to play TRIVIA?");
			Console.WriteLine("Press Enter to play...\n");
			while (Console.ReadKey().Key != ConsoleKey.Enter) { }


			while (true)
			{

				int columnPicker = rand.Next(colCount);
				while (columnPicker == 2 || columnPicker == 3 || columnPicker == 5 || columnPicker == 6 || columnPicker == 0)
				{
					columnPicker = rand.Next(colCount);
				}
				int rowPicker = rand.Next(rowCount);

				String question = xlRange.Cells[rowPicker, columnPicker].Value.ToString();
				String answer = xlRange.Cells[rowPicker, columnPicker + 1].Value.ToString();

				String answer2 = xlRange.Cells[rand.Next(rowCount), columnPicker + 1].Value.ToString();
				String answer3 = xlRange.Cells[rand.Next(rowCount), columnPicker + 1].Value.ToString();
				String answer4 = xlRange.Cells[rand.Next(rowCount), columnPicker + 1].Value.ToString();

				String[] answers = { answer, answer2, answer3, answer4 };
				var answersList = new List<String>(answers);

				Random orderPicker = new Random();


				Console.WriteLine("Question {0}: \n", i);

				Console.WriteLine(question);
				for (int j = 3; j >= 0; j--)
				{
					int order = orderPicker.Next(j);

					Console.WriteLine("{0}: {1}", j, answersList.ElementAt(order));
					answersList.RemoveAt(order);
				}

				Console.WriteLine("\nYour answer: ");
				String givenAnswer = Console.ReadLine().ToUpper();

				if (givenAnswer == answer)
				{
					userScore++;
					Console.WriteLine("\nCorrect!");
					Console.WriteLine("Your score is {0}", userScore);
					Console.WriteLine("Press any key to continue...\n");


				}
				else
				{
					userScore = 0;
					Console.WriteLine("\nIncorrect!");
					Console.WriteLine("The correct answer was {0}.", answer);
					Console.WriteLine("Your score is {0}", userScore);
					Console.WriteLine("Press any key to continue...\n");
				}
				Console.ReadKey();
				i++;
			}

		}

	}
}
