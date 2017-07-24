// MassiveAndPointer.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
using namespace std;

int main()
{
	setlocale(LC_ALL, "Russian");
	srand(2);
	const int width = 9;
	const int height = 5;
	char Lts[height][width];
	for (int i = 0;i < height;i++)
	{
		for (int j = 0;j < width;j++)
		{
			Lts[i][j] = 'A' + rand() % 25;
			cout <<Lts[i][j]<<" ";
		}
		cout << endl;
	}

	system("pause>nul");
    return 0;
}

