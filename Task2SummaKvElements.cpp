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
	const int size = 1;
	int summkv = 0;
	const int cols[size] = { 5};
	int** nums = new int*[size];
	for (i = 0;i < size;i++)
	{
		nums[i] = new int[cols[i]];
		cout << "| ";
		for (j = 0;j < cols[i];j++)
		{
				nums[i][j] = rand() % 10;
				cout << nums[i][j] << " | ";		
		}
		cout << endl;
	}


	for (j = 0;j<5;j++)
	{
		summkv += nums[0][j] * nums[0][j];
		
	}
	cout << "Сумма квадратов элементов равна = " << summkv << endl;

	for (i = 0;i < size;i++)
	{
		delete[] nums[i];
	}
	delete[] nums;

	system("pause>nul");
    return 0;
}

