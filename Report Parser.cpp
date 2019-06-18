/*
Online Courses: Export/Show/Writing C++ output to Excel, https://www.youtube.com/watch?v=CD-1reCKwHw
https://www.geeksforgeeks.org/split-a-sentence-into-words-in-cpp/

*/
#include <iostream>
#include <fstream>
#include <string>
#include <iomanip>
#include <sstream>
void postCards();
void fullReport();
using namespace std;

int main()
{
	postCards();	
	cout << "Formatted infomation to excel postcard template, press Enter to continue:\n";
	fullReport();
	cout << "Printed all information to an excel workbook, press Enter to close:\n";
	return 0;
}

void postCards()
{
	string info;
	ofstream outData;
	ifstream inData;
	
	inData.open("info.txt");
	outData.open("postcard.xls", ios::app);
	system("pause");

		while (getline(inData, info))
		{
			int count = 0;
			int iterate = 0;
			int space = iterate;

			for (int i = 0; i < info.length(); i++)
				if (info[i] == ' ')
					count++;

			int address = count - 10;
			istringstream ss(info);
			// Traverse through all words 
			do {

				if (iterate == address - 1)
					break;

				// Read a word 
				string word;
				ss >> word;
				//Erases key words from every row to line every column up
				if (word == "CN/Condo" || word == "SF" || word == "ACTV->Show" || word == "CN/Co-Op" || word == "Deposit->Show")
					continue;
				
				//Gets rid of the "N" on the 4th lines of every row
				if (iterate == 4 && word == "N")
					continue;
					
				//Combines names of towns that start with "New", "Old", etc.
				if (word == "New" || word == "Old" || word == "North" || word == "East" || word == "South" || word == "West" || word == "Beacon" || word == "Rocky" || word == "Deep" || word == "Windsor")
				{
						address--;
						outData << word << " ";
						continue;
				}

					iterate++;
				if (iterate <= 4)
					continue;
				
				outData << word << "\t";
			} while (ss);
			outData << "\n";
		}
		inData.close();
		outData.close();	
}

void fullReport()
{
	string info;
	ofstream outData;
	ifstream inData;
	inData.open("info.txt");
	outData.open("fullreport.xls", ios::app);
	system("pause");
		while (getline(inData, info))
		{
			int count = 0;
			
			int iterate = 0;
			for (int i = 0; i < info.length(); i++)
				if (info[i] == ' ')
					count++;
			int address = count - 10;
			int space = address;
			istringstream ss(info);
			// Traverse through all words 
			do {
				iterate++;
				// Read a word 
				string word;
				ss >> word;
				if (word == "N")
				{
					word.erase();
					continue;
				}
				//Starts a new cell after this keyword
				if (iterate == 3 && word != "ACTV->Show")
					outData << "\t";

				//  Counts and Combines towns that start with "New", "Old", etc.
				if (word == "New" || word == "Old" || word == "North" || word == "East" || word == "South" || word == "West" || word == "Beacon" || word == "Rocky" || word == "Deep")
				{
					outData << word << " ";
					continue;
				}
				
				// Print the read word 
				// While there is more to read 
				
				outData << word << "\t";
			} while (ss);
			outData << "\n";
		}
	inData.close();
	outData.close();
}
