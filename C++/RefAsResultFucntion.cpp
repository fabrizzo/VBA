// MassiveAndPointer.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <cstdio>
using namespace std;

int &getMax(int* nums, int n)
{
	int i = 0, k;
	for (k = 0;k < n;k++)
	{
		if (nums[k] > nums[i])
		{
			i = k;
		}
	}
	return nums[i];
}
void show(int* nums, int n)
{
	for (int i = 0;i < n;i++)
	{
		cout << nums[i] << " ";
	}
}
int main()
{
	setlocale(LC_ALL, "Russian");
	const int size = 10;
	int numbers[size] = { 1,5,8,2,4,9,12,3 };
	show(numbers, size);
	int maxNum = getMax(numbers, size);
	cout << "Максимальное значение: " << maxNum << endl;
	maxNum = -100;
	show(numbers, size);
	int &maxRef = getMax(numbers, size);
	cout << "Максимальное значение: " << maxRef << endl;
	maxRef = -200;
	show(numbers, size);
	cout << "Максимальное значение: ";
	cout << getMax(numbers, size) << endl;

	
	system("pause>nul");
	return 0;
}

