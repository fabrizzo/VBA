// MassiveAndPointer.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <cstdio>
using namespace std;
void show(int** M, int p, int n)
{
	for (int i = 0;i < p;i++)
	{
		for (int j = 0;j < n;j++)
		{
			printf("%4d", M[i][j]);
		}
		printf("\n");
	}
}
int main()
{
	setlocale(LC_ALL, "Russian");
	int a = 3, b = 5, i, j;
	int** A = new int*[a];
	for (i = 0;i < a;i++)
	{
		A[i] = new int[b];
		for (j = 0;j < b;j++)
		{
			A[i][j] = i*b + j + 1;
		}
	}
	printf("Содержимое массива: \n");
	show(A, a, b);
	for (i = 0;i < a;i++)
	{
		delete[] A[i];
	}
	delete[] A;
	system("pause>nul");
	return 0;
}

