// MassiveAndPointer.cpp: ���������� ����� ����� ��� ����������� ����������.
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
	int A[n][n] = { { 1,-2,1 },{ 2,0,-1 },{ 2,3,-1 } };
	printf("������� �:\n");
	show(A);
	int B[n][n];
	for (int i = 0;i < n;i++)
		{
		for (int j = 0;j < n;j++)
			{
				B[i][j] = 0;
			for (int k = 0;k < n;k++)
				{
					B[i][j] = A[j][i];
				}
			}
		}
	printf("������� �����������������:\n");
	show(B);


	system("pause>nul");
	return 0;
}

