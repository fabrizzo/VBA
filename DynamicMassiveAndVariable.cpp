// MassiveAndPointer.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
using namespace std;

int main()
{
	setlocale(LC_ALL, "Russian");
	int* size;
	size = new int;
	cout << "Введите размер массива: ";
	cin >> *size;
	char* symbs;
	symbs = new char[*size];
	for (int k = 0;k < *size;k++)
	{
		symbs[k] = 'a' + k;
		cout << symbs[k] << " ";
	}
	delete[] symbs;
	delete size;
	cout << "\nМассив и переменная удалены\n";




	system("pause>nul");
    return 0;
}

