#include"Excel.h"

HANDLE  hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
int main()
{
	SetConsoleTextAttribute(hConsole, 15);
	cout << "\n Desgin by Vu Duc Minh , Nguyen Trung Thinh , Diep Hai Trieu ";
	Excel b;
	string c;
	string option;
	SetConsoleTextAttribute(hConsole, 175);
	cout << "\n -------------------EXCEL PROGRAM-------------------------" << endl;  //nen xanh la,chu trang 175
	SetConsoleTextAttribute(hConsole, 15);
	cout << "\n -------------------USER MANUAL---------------------------";
	SetConsoleTextAttribute(hConsole, 4);
	cout << "\n Please input correctly. If not,it would be skipped ! ";


	SetConsoleTextAttribute(hConsole, 15);
	cout << "\n Press y to add sheets from text file ; press others to pass: ";
	rewind(stdin);
	getline(cin, option);

	transform(option.begin(), option.end(), option.begin(), toupper);
	if (option == "Y")
	{
		cout << "\n This is a default Excel" << endl;
		b.readSheet();
		system("cls");
		b.print();
	}
	do
	{
		string choice;
		cout << "\n Case 1: Choose if you want to input Sheets";
		cout << "\n Case 2: Choose if you want to update sheets";
		cout << "\n Case 3: Choose if you want to exit!!!!!";
		cout << "\n Choose your options: ";
		rewind(stdin);
		getline(cin, c);
		if (c == "1")
		{
			b.addSheet();
			b.print();
		}
		else if (c == "2")
		{
			b.update();
			system("cls");
			b.print();
		}
		else if (c == "3")
		{
			break;
		}
	} while (c != "3");
	system("cls");
	b.print();





	_getch();
	return 0;
}