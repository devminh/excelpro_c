#include "Cells.h"

void Cells::Input()
{
	cout << "\n Input cell: ";
	fflush(stdin);
	getline(cin, data);
}

void Cells::Output()
{
	cout << data;
}
