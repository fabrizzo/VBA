// MassiveAndPointer.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
using namespace std;

int main()
{
	setlocale(LC_ALL, "Russian");
	char str[100] = "Программируем на С++";
	cout << str << endl;
	for (int k = 0;str[k];k++)
	{
		cout << str[k] << "_";
	}
	cout << endl;
	for (char* p = str; *p;p++)
	{
		cout << p << endl;
	}
	str[13] = '\0';
	cout << str << endl;
	cout << str + 14 << endl;
	cout << "Раз два три" + 4 << endl;
	const char* q = "Раз два три" + 8;
	cout << q[0] << endl;
	cout << q << endl;

	system("pause>nul");
    return 0;
}

