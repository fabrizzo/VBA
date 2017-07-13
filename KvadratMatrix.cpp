// MassiveAndPointer.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <cstdio>
using namespace std;
const int n = 3;
void show(int M[n][n])
{
	for (int i = 0;i < n;i++)
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
	int A[n][n] = { {1,-2,1},{2,0,-1},{2,3,-1} };
	printf("Матрица А:\n");
	show(A);
	int B[n][n] = { {2,1,-1},{1,3,1},{-2,1,4} };
	printf("Матрица B:\n");
	show(B);
	int C[n][n];
	for (int i = 0;i < n;i++)
	{
		for (int j = 0;j < n;j++)
		{
			C[i][j] = 0;
			for (int k = 0;k < n;k++)
			{
				C[i][j] += A[i][j] * B[i][j];
			}
		}
	}
	printf("Матрица С=А*В:\n");
	show(C);


	system("pause>nul");
    return 0;
}

