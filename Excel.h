#pragma once
#include"Sheets.h"

class Excel
{
private:
	map<string, Sheets> ex;  //address,sheets
	
	string add;//address of sheet
public:

	void readSheet();
	void addSheet();
	
	string getSheetAddress();

	void update();

	void print();
	
	

};