#include "Sheets.h"

bool kt(string str)
{
	if (str == "")
		return false;



	int count1 = 0;
	for (int i = 0; i < str.length(); i++)
	{
	 

		if (str[i] == '.')
		{
			count1++;
			if (i == 0 || count1 == 2)
				return false;
			if (!isdigit(str[i + 1]) || !isdigit(str[i - 1]))
				return false;
		}
		else if (str[i] == '-'||str[i]=='+')
		{
			if (str[i + 1] == '-'||str[i+1] == '+')
			{
				if (isdigit(str[i + 2]))
					return true;
			}
			else if(isdigit(str[i+1]))
				return true;
			return false;
		}
		else if (!isdigit(str[i]))
			return false;
		
	}
	
	return true;
}

void Sheets::CreateSheet()
{
	
	cout << "\n Input the quantity of columns: ";
	 cin >> col;
	cout << "\n Input the quantity of rows: ";
	 cin >> row;
	 
	 string choice;
	 cout << "\n Do you want to select other cells ?: 1:Yes, press others to cancel ";
	 cin >> choice;
	 while (choice == "1")
	 {
	 
	
	 string s1;
	 cout << "\n Input the address :";
	 rewind(stdin);
	 getline(cin, s1);
     cout << "\n Input data: ";
	 rewind(stdin);
	 getline(cin, m[s1]);

	
	
	 if (kt(m[s1])==true) //check cells if they are numbers
	 {
		 m1[s1] = stod(m[s1]); //string to double
		 
		 
	 }
	 
	 cout << "\n Do you want to select other cells ?: 1:Yes, press others to cancel ";
	 cin >> choice;
	 
	 } 
	 




}

void Sheets::input()
{
	string choice;
	cout << "\n Press 1 to continue to update , press others to exit ! ";
	cin >> choice;
	while (choice == "1")
	{


		string s1;
		cout << "\n Input the address :";
		rewind(stdin);
		getline(cin, s1);

		cout << "\n Input data: ";
		rewind(stdin);
		getline(cin, m[s1]);



		if (kt(m[s1]) == true) //check cells if they are numbers
		{
			m1[s1] = stod(m[s1]); //string to double


		}

		cout << "\n Do you want to select other cells ?: 1:Yes,2:No ";
		cin >> choice;

	}


}

void Sheets::PrintSheet()
{
	HANDLE  hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
	string s;
	cout << "     ";
	for (int i = 0; i < col; i++)
	{
		SetConsoleTextAttribute(hConsole, 230);
		s = char(i + 65);
		cout << setw(9) << left << s;



	}
	
	for (int i = 0; i < row; i++)
	{
		SetConsoleTextAttribute(hConsole, 158);
		cout << endl;
		cout << "|";

		cout << setw(3) << left << i + 1;
		
		cout << "|";
		
		for (int j = 0; j < col; j++)//column
		{
			
			/*cout << t << endl;*/
			SetConsoleTextAttribute(hConsole, 240);
			s = char(j + 65) + to_string(i + 1);  //A 1 B 2 C 5 -> DOC DIA CHI

			if (m[s].size() <= 8)
					cout << setw(8) << left << m[s] << "|";
			else
			{
				string temp = m[s].substr(0, 5) + "...";
				cout << setw(8) << left << temp << "|";
			}
		}
		cout <<endl;
	}
	
	SetConsoleTextAttribute(hConsole, 15);
	cout << "\n There are "<<demo<<" address ; "<<"There are "<<demso<<" numbers ; "<<"There are: "<< demchuoi<<" string ; There are "<<demo + demso+demchuoi<<" data ";
	cout << "\n \n";
}

void Sheets::readFile()
{
	int c;
	string r;

	string data;
	string add;
	fstream fs;

	fs.open("Data.txt");
	while (!fs.eof())
	{
		getline(fs, r);
		c = r.find_first_of(" "); //find the first space 
		add = r.substr(0, c);
		
		r.erase(0, c + 1);
		data = r;

		m[add] = data;

		if (kt(m[add]) == true) //check cells if they are numbers
		{
			m1[add] = stod(m[add]); //string to double
		}
		
	}
	fs.close();
}

void Sheets::countInSheet()
{
	
	string s;
	string sub;
	string sub2; // quet ham count,quet ham sumif,aveif
	string sub3; //quet ham countif
	string sub4; //quet ham copy
	string sub5; //quet ham cut
	string sub6;
	int ran = 69;
	for (int i = 0; i < row; i++)
	{
		for (int j = 0; j < col; j++)//column
		{
			s = char(j + 65) + to_string(i + 1);  //A 1 B 2 C 5 -> DOC DIA CHI
			sub = m[s].substr(0, 5);  //take from = to '('

			sub2 = m[s].substr(0, 7);
			sub3 = m[s].substr(0, 9);
			sub4 = m[s].substr(0, 6);
			sub5 = m[s].substr(0, 5);
			sub6= m[s].substr(0, 10);
			transform(sub.begin(), sub.end(), sub.begin(), toupper);

			transform(sub2.begin(), sub2.end(), sub2.begin(), toupper);

			transform(sub3.begin(), sub3.end(), sub3.begin(), toupper);

			transform(sub4.begin(), sub4.end(), sub4.begin(), toupper);

			transform(sub5.begin(), sub5.end(), sub5.begin(), toupper);
			transform(sub6.begin(), sub6.end(), sub6.begin(), toupper);

			if (sub == "=SUM(")
			{
				demo++;
			}
			else if (sub == "=AVE(")
			{
				demo++;
				
			}
			else if (sub2 == "=COUNT(")
			{
				demo++;
			
			}
			else if (sub2 == "=SUMIF(")
			{
				demo++;
				
			}
			else if (sub2 == "=AVEIF(")
			{
				demo++;
				
			}
			else if (sub3 == "=COUNTIF(")
			{
				demo++;
				
			}
			else if (sub4 == "=COPY(")
			{
				demo++;
				
			}
			else if (sub5 == "=CUT(")
			{
				demo++;
			}
			else if (sub5 == "=AND(")
			{
				demo++;
			}
			else if (sub6 == "=AUTOFILL(")
			{
				demo++;
			}
			else if (sub3 == "=VLOOKUP(")
			{
				demo++;
			}
			else if (kt(m[s]))
			{
				demso++;


			}
			else if (!kt(m[s]) && m[s] != "" )
			{
				demchuoi++;
			}

		}

	}
	


}

double Sheets::sumInRange(int first_r, int last_r, int first_c, int last_c,int &count)
{
	double s = 0;
	string t;
	for (int i = first_r - 1; i < last_r; i++)  //from row i to row j
	{
		
		for (int j = first_c - 1; j < last_c; j++) //from col i to col j
		{  
			t = char(j + 65) + to_string(i + 1); //numbers of column to  A,B,C,D
			if (kt(m[t]))
			{
				
				count++;
				s += m1[t];
			}
		}
	}

	return s;
}

double Sheets::sum(string t,int &count, map<string, Sheets> ex, int dc, int dr)
{
	double sum = 0;
	transform(t.begin(), t.end(), t.begin(), toupper);  //change to upper
	string sub1;
	int  pos1;
	pos1 = t.find_first_of("(");

	t.erase(0, pos1 + 1);
	t.pop_back(); //delete final words
	
	while (t != "")
	{
		

		//meet : 
		int pos2;
		string sub2; string sub3;  // (a1)  delete (:)  (d3)  

		int pos3 = t.find_first_of(',');
		
		sub3 = t.substr(0, pos3);
		
		t.erase(0, pos3 + 1);
		if (pos3 == string::npos)
		{
			t = "";
		}

		if (sub3.find_first_of(':') != string::npos) //npos is a constant that shows letter is not exist
		{
			if (sub3.find_first_of("SHEET") != string::npos)
			{
				string sheetadd = sub3.substr(0, 6);   //take out Sheet1

				sub3.erase(0, sub3.find_first_of('!') + 1);
				pos2 = sub3.find_first_of(':');
				sub1 = sub3.substr(0, pos2); //address of left side

				sub3.erase(0, pos2 + 1);
				sub3.erase(0, sub3.find_first_of('!') + 1);
				sub2 = sub3;//address of right side

				int r1, r2, c1, c2;
				c1 = min(int(sub1[0]) - 64, int(sub2[0]) - 64);  //smallest column
				c2 = max(int(sub1[0]) - 64, int(sub2[0]) - 64);

				string a, b;
				a = sub1.erase(0, 1);
				b = sub2.erase(0, 1);


				r1 = min(stoi(a), stoi(b));
				r2 = max(stoi(a), stoi(b));

				sum += ex[sheetadd].sumInRange(r1+dr, r2+dr, c1+dc, c2+dc, count);
			}
			else
			{
                pos2 = sub3.find_first_of(':');
				sub1 = sub3.substr(0, pos2); //address of left side

				sub3.erase(0, pos2 + 1);

				pos2 = sub3.find_first_of(',');
				sub2 = sub3.substr(0, pos2);  //address of right side
				sub3.erase(0, pos2 + 1);

				int r1, r2, c1, c2;
				c1 = min(int(sub1[0]) - 64, int(sub2[0]) - 64);  //smallest column
				c2 = max(int(sub1[0]) - 64, int(sub2[0]) - 64);

				string a, b;
				a = sub1.erase(0, 1);
				b = sub2.erase(0, 1);


				r1 = min(stoi(a), stoi(b));
				r2 = max(stoi(a), stoi(b));

				sum += sumInRange(r1+dr, r2+dr, c1+dc, c2+dc, count);
			}
		}
		else
		{
			if (sub3.find_first_of("SHEET") != string::npos)
			{
				string sheetadd = sub3.substr(0, 6);   //take out Sheet1

				sub3.erase(0, sub3.find_first_of('!') + 1);

				//
				string t = sub3.substr(0, 1);

				int cc = int(sub3[0]) + dc;

				sub3.erase(0, 1);
				int rr = stoi(sub3) + dr;
				
				
				sub3 = char(cc) + to_string(rr);
				//

				if (kt(ex[sheetadd].m[sub3]))
				{
					sum += ex[sheetadd].m1[sub3];
					count++;
				}
			}
			else
			{
				//
				string t = sub3.substr(0, 1);

				int cc = int(sub3[0]) + dc;

				sub3.erase(0, 1);
				int rr = stoi(sub3) + dr;


				sub3 = char(cc) + to_string(rr);
				//

				if (kt(m[sub3]))
				{
					sum += m1[sub3];
					count++;
				}
			}
		}
		
	}
	return sum;
}

double Sheets::sumInRangeIf(int first_r, int last_r, int first_c, int last_c, int & count, double dks, string dkd)
{

    double s = 0;
	string t;
	for (int i = first_r - 1; i < last_r; i++)  //from row i to row j
	{

		for (int j = first_c - 1; j < last_c; j++) //from col i to col j
		{
			t = char(j + 65) + to_string(i + 1); //numbers of column to  A,B,C,D
			if (kt(m[t]))
			{
				if (dkd == ">")
				{
					if (m1[t] > dks)
					{
						count++;
						s += m1[t];
					}


				}
				else if (dkd == "<")
				{
					if (m1[t] < dks)
					{
						count++;
						s += m1[t];
					}


				}
				else if (dkd == "!=")
				{
					if (m1[t] != dks)
					{
						count++;
						s += m1[t];
					}


				}

			
			}
		}
	}

	return s;



}

double Sheets::sumIF(string t, int & count, map<string, Sheets> ex, int dc, int dr)
{
	double sum = 0;
	transform(t.begin(), t.end(), t.begin(), toupper);  //change to upper
	
	string sub1;
	int  pos1;
	pos1 = t.find_first_of("(");

	t.erase(0, pos1 + 1);
	t.pop_back(); //delete final words

	string dk; //take out the condition
	int temp = t.find_first_of("'");

	int temp1 = t.find_last_of("'");
	string dkd;
	double dks;
	dk = t.substr(temp, temp1);
    dk.erase(0, 1);
	dk.pop_back();
	

	if (dk.find_first_of('!') != string::npos)
	{
		dkd= dk.substr(0, 2);  //lay dau !=
		dk.erase(0, 2);
		 dks = stod(dk);

	}
	else
	{
		dkd = dk.substr(0, 1);  //lay dau

		dk.erase(0, 1);

		 dks = stod(dk);   //lay so
	}
	t.erase(t.find_last_of(','), t.size()-1);  //this is solved
    while (t != "")
	{
		//meet : 
		int pos2;
		string sub2; string sub3;   

		int pos3 = t.find_first_of(',');

		sub3 = t.substr(0, pos3);  // a1:b3

		t.erase(0, pos3 + 1);
		if (pos3 == string::npos)
		{
			t = "";
		}

		if (sub3.find_first_of(':') != string::npos) //npos is a constant that shows letter is not exist
		{
			if (sub3.find_first_of("SHEET") != string::npos)
			{
				string sheetadd = sub3.substr(0, 6);   //take out Sheet1

				sub3.erase(0, sub3.find_first_of('!') + 1);
				pos2 = sub3.find_first_of(':');
				sub1 = sub3.substr(0, pos2); //address of left side

				sub3.erase(0, pos2 + 1);
				sub3.erase(0, sub3.find_first_of('!') + 1);
				sub2 = sub3;//address of right side

				int r1, r2, c1, c2;
				c1 = min(int(sub1[0]) - 64, int(sub2[0]) - 64);  //smallest column
				c2 = max(int(sub1[0]) - 64, int(sub2[0]) - 64);

				string a, b;
				a = sub1.erase(0, 1);
				b = sub2.erase(0, 1);


				r1 = min(stoi(a), stoi(b));
				r2 = max(stoi(a), stoi(b));

				sum += ex[sheetadd].sumInRangeIf(r1+dr, r2+dr, c1+dc, c2+dc, count,dks,dkd);
			}
			else
			{
				pos2 = sub3.find_first_of(':');
				sub1 = sub3.substr(0, pos2); //address of left side

				sub3.erase(0, pos2 + 1);

				pos2 = sub3.find_first_of(',');
				sub2 = sub3.substr(0, pos2);  //address of right side
				sub3.erase(0, pos2 + 1);

				int r1, r2, c1, c2;
				c1 = min(int(sub1[0]) - 64, int(sub2[0]) - 64);  //smallest column
				c2 = max(int(sub1[0]) - 64, int(sub2[0]) - 64);

				string a, b;
				a = sub1.erase(0, 1);
				b = sub2.erase(0, 1);


				r1 = min(stoi(a), stoi(b));
				r2 = max(stoi(a), stoi(b));

				sum += sumInRangeIf(r1+dr, r2+dr, c1+dc, c2+dc, count, dks, dkd);
			}
		}
		else
		{
			if (sub3.find_first_of("SHEET") != string::npos)
			{
				string sheetadd = sub3.substr(0, 6);   //take out Sheet1
                sub3.erase(0, sub3.find_first_of('!') + 1);
				
				string t = sub3.substr(0, 1);

				int cc = int(sub3[0]) + dc;

				sub3.erase(0, 1);
				int rr = stoi(sub3) + dr;


				sub3 = char(cc) + to_string(rr);


				if (kt( ex[sheetadd].m[sub3]) )
				{
					if (dkd == ">")
					{
						if (ex[sheetadd].m1[sub3] > dks)
						{
							count++;
							sum += ex[sheetadd].m1[sub3];
						}

                    }
					else if (dkd == "<")
					{
						if (ex[sheetadd].m1[sub3] < dks)
						{
							count++;
							sum += ex[sheetadd].m1[sub3];
						}


					}
					else if (dkd == "!=")
					{
						if (ex[sheetadd].m1[sub3] != dks)
						{
							count++;
							sum += ex[sheetadd].m1[sub3];
						}
                    }
				}
			}
			//
			else
			{
				string t = sub3.substr(0, 1);

				int cc = int(sub3[0]) + dc;

				sub3.erase(0, 1);
				int rr = stoi(sub3) + dr;


				sub3 = char(cc) + to_string(rr);

				//
				if (kt(m[sub3]))
				{
					if (dkd == ">")
					{
						if (m1[sub3] > dks)
						{
							count++;
							sum += m1[sub3];
						}
					}
					else if (dkd == "<")
					{
						if (m1[sub3] < dks)
						{
							count++;
							sum += m1[sub3];
						}


					}
					else if (dkd == "!=")
					{
						if (m1[sub3] != dks)
						{
							count++;
							sum += m1[sub3];
						}


					}
				}
			}
		}

	}
	return sum;
}

void Sheets::scanSheet(map<string, Sheets> &ex)
{
	double tong;
	double avg;
	string s;
	string sub;
	string sub2; // quet ham count,quet ham sumif,aveif
	string sub3; //quet ham countif,vlookup
	string sub4; //quet ham copy
	string sub5; //quet ham cut
	string sub6;
	int ran = 69;
	try
	{
		for (int i = 0; i < row; i++)
		{
			for (int j = 0; j < col; j++)//column
			{
				s = char(j + 65) + to_string(i + 1);  //A 1 B 2 C 5 -> DOC DIA CHI
				sub = m[s].substr(0, 5);  //take from = to '('

				sub2 = m[s].substr(0, 7);
				sub3 = m[s].substr(0, 9);
				sub4 = m[s].substr(0, 6);
				sub5 = m[s].substr(0, 5);
				sub6 = m[s].substr(0, 10);
				transform(sub.begin(), sub.end(), sub.begin(), toupper);

				transform(sub2.begin(), sub2.end(), sub2.begin(), toupper);

				transform(sub3.begin(), sub3.end(), sub3.begin(), toupper);

				transform(sub4.begin(), sub4.end(), sub4.begin(), toupper);

				transform(sub5.begin(), sub5.end(), sub5.begin(), toupper);

				transform(sub6.begin(), sub6.end(), sub6.begin(), toupper);

				string zero = "0";
				string zero1 = "0.000000";
				if (m[s] == zero || m[s] == zero1)
				{
					m[s] = "";
				}
				else if (sub == "=SUM(")
				{
					tong = this->sum(m[s], ran, ex, 0, 0);
					m[s] = to_string(tong);


				}
				else if (sub == "=AVE(")
				{
					avg = this->AVE(m[s], ex, 0, 0);
					m[s] = to_string(avg);


				}
				else if (sub2 == "=COUNT(")
				{
					m[s] = to_string(this->count(m[s], ex, 0, 0));


				}
				else if (sub2 == "=SUMIF(")
				{
					m[s] = to_string(this->sumIF(m[s], ran, ex, 0, 0));

				}
				else if (sub2 == "=AVEIF(")
				{
					m[s] = to_string(this->aveIF(m[s], ex, 0, 0));


				}
				else if (sub3 == "=COUNTIF(")
				{
					m[s] = to_string(this->countIF(m[s], ex, 0, 0));


				}
				else if (sub4 == "=COPY(")
				{
					m[s]=this->copy(m[s], ex, s, 0, 0);
				}
				else if (sub5 == "=CUT(")
				{
					m[s] = this->cut(m[s], ex, 0, 0);

				}
				else if (sub5 == "=AND(")
				{
					m[s] = this->AND(m[s], ex, 0, 0);


				}
				else if (sub3 == "=VLOOKUP(")
				{
					m[s] = this->cutStr(m[s],ex);
				}


			}
			
		}
		throw std::runtime_error("");
	}
	catch (std::exception const& e)
	{
		cout << "";
	}

}

double Sheets::AVE(string t,map<string, Sheets> ex, int dc, int dr)
{
	int count = 0;
	double s = this->sum(t,count,ex,dc,dr);
	if (count == 0)
		count = 1;
	return s / count;
}

int Sheets::count(string t, map<string, Sheets> ex, int dc, int dr)
{
	int dem = 0;
	this->sum(t, dem,ex,dc,dr);
	
	return dem;
}

double Sheets::aveIF(string t, map<string, Sheets> ex, int dc, int dr)
{
	int count = 0;
	double s = this->sumIF(t, count, ex,dc,dr);
	if (count == 0)
		count = 1;
	return s / count;
}

int Sheets::countIF(string t, map<string, Sheets> ex, int dc, int dr)
{
	int dem = 0;
	this->sumIF(t, dem, ex,dc,dr);

	return dem;
}
 
void Sheets::addCells(string t)  
{
	string sub1; string sub2;
	int c;
	string add;
	string data;
	c = t.find_first_of(" "); //find the first space 
	
	add = t.substr(0, c); 



	t.erase(0, c + 1);
	data = t;

    m[add] = data;

	if (kt(m[add]) == true) //check cells if they are numbers
	{
		m1[add] = stod(m[add]); //string to double
	}
	
	sub1 = add;
	
	int cot = int(sub1[0]) - 64;

	sub1.erase(0, 1);
	
	
	int dong;
	dong = stoi(sub1);


	if (cot > col)
	{
		col = cot;
	}
	if (dong > row)
	{
		row = dong;
	}


}

string Sheets::andInRange(int first_r, int last_r, int first_c, int last_c)
{
	string s;
	string t;
	for (int i = first_r - 1; i < last_r; i++)  //from row i to row j
	{

		for (int j = first_c - 1; j < last_c; j++) //from col i to col j
		{
			t = char(j + 65) + to_string(i + 1); //numbers of column to  A,B,C,D
			if (kt(m[t])==false)
			{
                   s += m[t];
			}
		}
	}

	return s;
}

string Sheets::AND(string t, map<string, Sheets> ex, int dc, int dr)
{
	string sum;
	transform(t.begin(), t.end(), t.begin(), toupper);  //change to upper
	string sub1;
	int  pos1;
	pos1 = t.find_first_of("(");

	t.erase(0, pos1 + 1);
	t.pop_back(); //delete final words

	while (t != "")
	{


		//meet : 
		int pos2;
		string sub2; string sub3;  // (a1)  delete (:)  (d3)  

		int pos3 = t.find_first_of(',');

		sub3 = t.substr(0, pos3);

		t.erase(0, pos3 + 1);
		if (pos3 == string::npos)
		{
			t = "";
		}

		if (sub3.find_first_of(':') != string::npos) //npos is a constant that shows letter is not exist
		{
			if (sub3.find_first_of("SHEET") != string::npos)
			{
				string sheetadd = sub3.substr(0, 6);   //take out Sheet1

				sub3.erase(0, sub3.find_first_of('!') + 1);
				pos2 = sub3.find_first_of(':');
				sub1 = sub3.substr(0, pos2); //address of left side

				sub3.erase(0, pos2 + 1);
				sub3.erase(0, sub3.find_first_of('!') + 1);
				sub2 = sub3;//address of right side

				int r1, r2, c1, c2;
				c1 = min(int(sub1[0]) - 64, int(sub2[0]) - 64);  //smallest column
				c2 = max(int(sub1[0]) - 64, int(sub2[0]) - 64);

				string a, b;
				a = sub1.erase(0, 1);
				b = sub2.erase(0, 1);

				r1 = min(stoi(a), stoi(b));
				r2 = max(stoi(a), stoi(b));
				

				sum += ex[sheetadd].andInRange(r1+dr, r2+dr, c1+dc, c2+dc);
			}
			else
			{
				pos2 = sub3.find_first_of(':');
				sub1 = sub3.substr(0, pos2); //address of left side

				sub3.erase(0, pos2 + 1);

				pos2 = sub3.find_first_of(',');
				sub2 = sub3.substr(0, pos2);  //address of right side
				sub3.erase(0, pos2 + 1);

				int r1, r2, c1, c2;
				c1 = min(int(sub1[0]) - 64, int(sub2[0]) - 64);  //smallest column
				c2 = max(int(sub1[0]) - 64, int(sub2[0]) - 64);

				string a, b;
				a = sub1.erase(0, 1);
				b = sub2.erase(0, 1);


				r1 = min(stoi(a), stoi(b));
				r2 = max(stoi(a), stoi(b));

				sum += andInRange(r1 + dr, r2 + dr, c1 + dc, c2 + dc);
			}
		}
		else
		{
			if (sub3.find_first_of("SHEET") != string::npos)
			{
				string sheetadd = sub3.substr(0, 6);   //take out Sheet1

				sub3.erase(0, sub3.find_first_of('!') + 1);
				if (kt(ex[sheetadd].m[sub3])==false)
				{
					sum += ex[sheetadd].m[sub3];
					
				}
			}

			if (kt(m[sub3])==false)
			{
				sum += m[sub3];
				
			}
		}

	}

	return sum;


}

string Sheets::copy(string add, map<string, Sheets> ex , string curr, int dc, int dr)   //chua dung den sheet
{
	string data;
	
	int pos1 = add.find_first_of('(');
	add.erase(0, pos1 + 1);
	add.pop_back();

	transform(add.begin(), add.end(), add.begin(), toupper);


	if (add.find_first_of("SHEET") != string::npos)
	{
		string sheetadd = add.substr(0, 6);   //take out Sheet1

		add.erase(0, add.find_first_of('!') + 1);
		

        
		string t = add.substr(0, 1);

		int cc = int(add[0]) + dc;

		add.erase(0, 1);
		int rr = stoi(add) + dr;


		add = char(cc) + to_string(rr);
		


		data = ex[sheetadd].m[add];

		if (kt(ex[sheetadd].m[add]) == true)
		{
			m1[curr] = stod(ex[sheetadd].m[add]);


		}
	}

	else
	{
		
		string t = add.substr(0, 1);

		int cc = int(add[0]) + dc;

		add.erase(0, 1);
		int rr = stoi(add) + dr;


		add = char(cc) + to_string(rr);

		data = m[add];
		 

		 if (kt(m[add]) == true)
		 {
			m1[curr] = stod(m[add]);
			
	
		 }
		 
	}
	return data;
}

string Sheets::cut(string add, map<string, Sheets> &ex, int dc, int dr)
{
	string data;
	int pos1 = add.find_first_of('(');
	add.erase(0, pos1 + 1);
	add.pop_back();
	transform(add.begin(), add.end(), add.begin(), toupper);
	if (add.find_first_of("SHEET") != string::npos)
	{
		string sheetadd = add.substr(0, 6);   //take out Sheet1

		add.erase(0, add.find_first_of('!') + 1);


		string t = add.substr(0, 1);

		int cc = int(add[0]) + dc;

		add.erase(0, 1);
		int rr = stoi(add) + dr;


		add = char(cc) + to_string(rr);


		

		data = ex[sheetadd].m[add];
		ex[sheetadd].m[add] = "";
	}

	else
	{
		string t = add.substr(0, 1);

		int cc = int(add[0]) + dc;

		add.erase(0, 1);
		int rr = stoi(add) + dr;


		add = char(cc) + to_string(rr);

		data = m[add];
		m[add] = "";
	}
	return data;
}

void Sheets::autoFill(string t, string curr, map<string, Sheets> &ex)
{
	
	transform(t.begin(), t.end(), t.begin(), toupper);  //change to upper
	transform(curr.begin(), curr.end(), curr.begin(), toupper);
	//tach dia chi curr
	string before = curr;
	string temp = curr.substr(0, 1);
	int c1 = (temp[0]) - 64;
	curr.erase(0, 1);
	int r1 = stoi(curr);



	string sub1;
	int  pos1;
	pos1 = t.find_first_of("(");

	t.erase(0, pos1 + 1);
	t.pop_back(); //delete final words

	 pos1 = t.find_first_of(',');

	 string s = t.substr(0, pos1);
	string add = t.substr(0, pos1);
	string temp1 = add.substr(0, 1);
	
	int c2= (temp1[0]) - 64;
	add.erase(0, 1);
	int r2 = stoi(add);


	t.erase(0, pos1 + 1);
	

	int dc = c1 - c2; //distance of column
	int dr = r1 - r2; //distance of row

	if (t == "FALSE")
	{

		double tong;
		double avg;

		string sub;
		string sub2; // quet ham count,quet ham sumif,aveif
		string sub3; //quet ham countif
		string sub4; //quet ham copy
		string sub5; //quet ham cut
		int ran = 69;

		sub = m[s].substr(0, 5);  //take from = to '('
		sub2 = m[s].substr(0, 7);
		sub3 = m[s].substr(0, 9);
		sub4 = m[s].substr(0, 6);
		sub5 = m[s].substr(0, 5);
		transform(sub.begin(), sub.end(), sub.begin(), toupper);

		transform(sub2.begin(), sub2.end(), sub2.begin(), toupper);

		transform(sub3.begin(), sub3.end(), sub3.begin(), toupper);

		transform(sub4.begin(), sub4.end(), sub4.begin(), toupper);

		transform(sub5.begin(), sub5.end(), sub5.begin(), toupper);

		if (sub == "=SUM(")
		{
			tong = this->sum(m[s], ran, ex, 0, 0);
			m[before] = to_string(tong);
		}
		else if (sub == "=AVE(")
		{
			avg = this->AVE(m[s], ex, 0, 0);
			m[before] = to_string(avg);
		}
		else if (sub2 == "=COUNT(")
		{
			m[before] = to_string(this->count(m[s], ex, 0, 0));
		}
		else if (sub2 == "=SUMIF(")
		{
			m[before] = to_string(this->sumIF(m[s], ran, ex, 0, 0));
		}
		else if (sub2 == "=AVEIF(")
		{
			m[before] = to_string(this->aveIF(m[s], ex, 0, 0));
		}
		else if (sub3 == "=COUNTIF(")
		{
			m[before] = to_string(this->countIF(m[s], ex, 0, 0));
		}
		else if (sub4 == "=COPY(")
		{
			m[before] = this->copy(m[s], ex, s, 0, 0);
		}
		else if (sub5 == "=CUT(")
		{
			m[before] = this->cut(m[s], ex, 0, 0);
		}
		else if (sub5 == "=AND(")
		{
			m[before] = this->AND(m[s], ex, 0, 0);
		}
	}

	else if (t=="TRUE")
	{
		
		double tong;
		double avg;
		
		string sub;
		string sub2; // quet ham count,quet ham sumif,aveif
		string sub3; //quet ham countif
		string sub4; //quet ham copy
		string sub5; //quet ham cut
		int ran = 69;

		sub = m[s].substr(0, 5);  //take from = to '('
        sub2 = m[s].substr(0, 7);
		sub3 = m[s].substr(0, 9);
		sub4 = m[s].substr(0, 6);
		sub5 = m[s].substr(0, 5);
		transform(sub.begin(), sub.end(), sub.begin(), toupper);

		transform(sub2.begin(), sub2.end(), sub2.begin(), toupper);

		transform(sub3.begin(), sub3.end(), sub3.begin(), toupper);

		transform(sub4.begin(), sub4.end(), sub4.begin(), toupper);

		transform(sub5.begin(), sub5.end(), sub5.begin(), toupper);

		if (sub == "=SUM(")
		{
			tong = this->sum(m[s], ran, ex, dc, dr);
			m[before] = to_string(tong);
		}
		else if (sub == "=AVE(")
		{
			avg = this->AVE(m[s], ex, dc, dr);
			m[before] = to_string(avg);
		}
		else if (sub2 == "=COUNT(")
		{
			m[before] = to_string(this->count(m[s], ex, dc, dr));
		}
		else if (sub2 == "=SUMIF(")
		{
			m[before] = to_string(this->sumIF(m[s], ran, ex, dc, dr));
		}
		else if (sub2 == "=AVEIF(")
		{
			m[before] = to_string(this->aveIF(m[s], ex, dc, dr));
		}
		else if (sub3 == "=COUNTIF(")
		{
			m[before] = to_string(this->countIF(m[s], ex, dc, dr));
		}
		else if (sub4 == "=COPY(")
		{
			m[before] = this->copy(m[s], ex, s, dc, dr);
		}
		else if (sub5 == "=CUT(")
		{
			m[before] = this->cut(m[s], ex, dc, dr);
		}
		else if (sub5 == "=AND(")
		{
			m[before] = this->AND(m[s], ex, dc, dr);
		}
	}


}

void Sheets::fillScan(map<string, Sheets>& ex)
{
	string s;
	string sub;
	string sub1;
	try
	{
		for (int i = 0; i < row; i++)
		{
			for (int j = 0; j < col; j++)//column
			{
				s = char(j + 65) + to_string(i + 1);  //A 1 B 2 C 5 -> DOC DIA CHI
				sub = m[s].substr(0, 10);  //take from = to '('
				transform(sub.begin(), sub.end(), sub.begin(), toupper);
				if (sub == "=AUTOFILL(")
				{
					this->autoFill(m[s], s, ex);
					
					
				}
				

			}
		}
		throw std::runtime_error("");
	}

	catch (std::exception const& e)
	{
		cout << "";
	}



}

string Sheets::cutStr(string s, map<string, Sheets> ex)
{
	s.erase(0,s.find_first_of('(')+1);
	s.pop_back();
	string find= s.substr(0,s.find_first_of(','));
	find.pop_back();
	find.erase(0, 1);
	s.erase(0, s.find_first_of(',') + 1);
	
	string range= s.substr(0, s.find_first_of(','));
	s.erase(0, s.find_first_of(',') + 1);
     
	int pos_c = stoi(s);

	return this->vLookUp(find, range, pos_c,ex);
	


}

string Sheets::vLookUp(string find, string range, int pos_c, map<string, Sheets> ex)
{
	int pos2;
	string sub1;
	string sub2;
	string kq = "Not found!";
	transform(range.begin(), range.end(), range.begin(), toupper);
	if (range.find_first_of("SHEET") != string::npos)
	{
		
		string sheetadd = range.substr(0, 6);   //take out Sheet1

		range.erase(0, range.find_first_of('!') + 1);
		pos2 = range.find_first_of(':');
		sub1 = range.substr(0, pos2); //address of left side

		range.erase(0, pos2 + 1);
		range.erase(0, range.find_first_of('!') + 1);
		sub2 = range;//address of right side

		int r1, r2, c1, c2;
		c1 = min(int(sub1[0]) - 64, int(sub2[0]) - 64);  //smallest column
		c2 = max(int(sub1[0]) - 64, int(sub2[0]) - 64);

		string a, b;
		a = sub1.erase(0, 1);
		b = sub2.erase(0, 1);


		r1 = min(stoi(a), stoi(b));
		r2 = max(stoi(a), stoi(b));
		string t;

		for (int i = r1 - 1; i < r2; i++)  //from row i to row j
		{
			t = char(c1 - 1 + 65) + to_string(i + 1); //numbers of column to  A,B,C,D
			if (find == ex[sheetadd].m[t])
			{

				kq = ex[sheetadd].m[char(c1 - 1 + pos_c + 64) + to_string(i + 1)];
				break;
			}


		}
		
	}

	else
	{
		pos2 = range.find_first_of(':');
		sub1 = range.substr(0, pos2); //address of left side

		range.erase(0, pos2 + 1);

		pos2 = range.find_first_of(',');
		sub2 = range.substr(0, pos2);  //address of right side
		range.erase(0, pos2 + 1);

	int r1, r2, c1, c2;
	c1 = min(int(sub1[0]) - 64, int(sub2[0]) - 64);  //smallest column
	c2 = max(int(sub1[0]) - 64, int(sub2[0]) - 64);

	string a, b;
	a = sub1.erase(0, 1);
	b = sub2.erase(0, 1);


	r1 = min(stoi(a), stoi(b));
	r2 = max(stoi(a), stoi(b));
	if (c1 + pos_c - 1 > c2)
	{
		return "Fail!";
	}


	 
	string t;
	for (int i = r1 - 1; i < r2; i++)  //from row i to row j
	{
		t = char(c1 - 1 + 65) + to_string(i + 1); //numbers of column to  A,B,C,D
		if (find == m[t])
		{

			kq = m[char(c1 - 1 + pos_c + 64) + to_string(i + 1)];
			break;
		}


	}

	}
	return kq;

}

