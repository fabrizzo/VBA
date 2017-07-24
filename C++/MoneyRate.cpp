// MassiveAndPointer.cpp: ���������� ����� ����� ��� ����������� ����������.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <cstdio>
using namespace std;

unsigned long getMoney(unsigned long m, double r)
{
	return m*(1 + r / 100);
}
unsigned long getMoney(unsigned long m, double r, int y)
{
	double s = m;
	for (int k = 1; k <= y;k++)
	{
		s *= (1 + r / 100);
	}
	return s;
}
unsigned long getMoney(unsigned long m, double r, int y, int n)
{
	return getMoney(m, r / n, y*n);
}
int main()
{
	setlocale(LC_ALL, "Russian");
	unsigned long money;
	double rate;
	cout << "������� ����� ������ :";
	cin >> money;
	cout << "������� ���������� ������ : ";
	cin >> rate;
	cout << "��������� �����: " << money << endl;
	cout << "������� ������: " << rate << endl;
	cout << "����� �� ���� ���: " << getMoney(money,rate) << endl;
	cout << "����� �� ���� ���(����������� 4 ���� � ���): " << getMoney(money, rate, 7, 4) << endl;
	cout << "����� �� 3 ���: " << getMoney(money, rate, 3) << endl;
	cout << "����� �� 3 ��� (���������� 4 ���� � ���): " << getMoney(money, rate, 3, 4) << endl;
	cout << "����� �� 5 ���: " << getMoney(money, rate, 5) << endl;
	cout << "����� �� 5 ��� (���������� 4 ���� � ���): " << getMoney(money, rate, 5, 4) << endl;
	cout << "����� �� 7 ���: " << getMoney(money,rate,7) << endl;
	cout << "����� �� 7 ��� (���������� 4 ���� � ���): " << getMoney(money,rate,7,4) << endl;
	cout << "����� �� 10 ���: " << getMoney(money, rate, 10) << endl;
	cout << "����� �� 10 ��� (���������� 4 ���� � ���): " << getMoney(money, rate, 10, 4) << endl;
	system("pause>nul");
	return 0;
}

