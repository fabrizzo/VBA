// FunctionPrint.cpp: определяет точку входа для консольного приложения.
//

#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <string>
using namespace std;
void println(int n);
void println(char y);
void println(float x);
void println(string &d);
int main()
{
	setlocale(LC_ALL, "Russian");
	int a;
	float b;
	char c;
	string d;
	a = 2;
	b = 3.5;
	c = 'X';
	d = "Hello";
	println(a);
	println(b);
	println(c);
	println(d);


	system("pause>nul");
    return 0;
}

void println(int e)
{
	cout << e << "\n";
}
void println(float x)
{
	cout << x << "\n";
}
void println(char y)
{
	cout << y << "\n";
}
void println(string &d)
{
	string* p;
	p = &d;
	for (int i = 0; i < sizeof(p); i++)
	{
		cout << p[i] << "\n";
	}
}
		