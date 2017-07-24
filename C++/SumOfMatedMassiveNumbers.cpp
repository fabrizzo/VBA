// MassiveAndPointer.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <cstdio>
using namespace std;

int main()
{
	setlocale(LC_ALL, "Russian");
	srand(2);
	int i, j;
	const int size = 3;
	const int cols[size] = { 4,4,4 };
	int** nums = new int*[size];
	for (i = 0;i < size;i++)
	{
		nums[i] = new int[cols[i]];
		cout << "| ";
		for (j = 0;j < cols[i];j++)
		{
			if (i == 2)
			{
				nums[i][j] = nums[i - 1][j]*nums[i-2][j];
				cout << nums[i][j] << " | ";

			}
			else if (i < 2)
			{
				nums[i][j] = rand() % 10;
				cout << nums[i][j] << " | ";
			}
		}
		cout << endl;
	}
	
	for (i = 0;i < size;i++)
	{
		delete[] nums[i];
	}
	delete[] nums;

	system("pause>nul");
    return 0;
}

