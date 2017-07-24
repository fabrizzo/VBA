// MassiveAndPointer.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
using namespace std;

int main()
{
	setlocale(LC_ALL, "Russian");
	int num = 100;
	int &ref = num;
	cout << num << "= num | ref = " << ref << endl;
	ref = 200;
	cout << num << "= num | ref = " << ref << endl;
	system("pause>nul");
    return 0;
}

