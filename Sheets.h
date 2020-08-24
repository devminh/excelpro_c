#pragma once
#include<iostream>
#include <ctype.h>
#include<fstream>
#include<Windows.h>
#include<stdlib.h>
#include<algorithm>
#include <string>
#include<map>
#include<iomanip>
#include<conio.h>
using namespace std;
//Nen dung Map

class Sheets
{
private:
	map<string, string> m;//address,data
	map<string, double> m1; //adress,double
	int col;
	int row;
	int demchuoi = 0, demso = 0, demo = 0;

public:
	
	void CreateSheet();
	void input();  //this function is used to update
	void PrintSheet();

	
	void readFile();
	
	void countInSheet();

	//CALCULATING FUNCTION
	double sumInRange(int first_r, int last_r, int first_c, int last_c,int &count); 
	
	double sum(string t, int &count, map<string, Sheets> ex,int dc,int dr);
	
	double sumInRangeIf(int first_r, int last_r, int first_c, int last_c, int &count,double dks,string dkd); //>20
	double sumIF(string t, int &count, map<string, Sheets> ex,int dc,int dr);


	void scanSheet(map<string, Sheets> &ex);


	double AVE(string t, map<string, Sheets> ex,int dc,int dr);

	int count(string t, map<string, Sheets> ex,int dc,int dr);  //count digits

	double aveIF(string t, map<string, Sheets> ex,int dc, int dr);
	int countIF(string t, map<string, Sheets> ex, int dc, int dr);

	void addCells(string t);


	string andInRange(int first_r, int last_r, int first_c, int last_c);
	string AND(string t, map<string, Sheets> ex, int dc, int dr);

	//copy and cut function
	string copy(string add, map<string, Sheets> ex, string curr, int dc, int dr);
	string cut(string add, map<string, Sheets> &ex, int dc, int dr);


	void autoFill(string t, string curr, map<string, Sheets> &ex);
	void fillScan(map<string, Sheets> &ex);

	string cutStr(string s, map<string, Sheets> ex);
	string vLookUp(string find,string range,int pos_c, map<string, Sheets> ex);

};