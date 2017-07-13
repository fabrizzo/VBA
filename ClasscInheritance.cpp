// MassiveAndPointer.cpp: ���������� ����� ����� ��� ����������� ����������.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <cstdio>
#include <string>
using namespace std;
class myMoney
{
public:
	string name;
	double money;
	double rate;
	int time;
	double getMoney()
	{
		double s = money;
		for (int k = 1;k <= time;k++)
		{
			s *= (1 + rate / 100);
		}
		return s;
	}
	void showAll()
	{
		cout << "���: " << name << endl;
		cout << "�����: " << money << endl;
		cout << "������ (%): " << rate << endl;
		cout << "������ (���): " << time << endl;
		cout << "�������� �����: " << getMoney() << endl;
		for (int k = 1;k <= 35;k++)
		{
			cout << "-";
		}
		cout << endl;

	}
	void setAll(string n, double m, double r, int t)
	{
		name = n;
		money = m;
		rate = r;
		time = t;
	}

	myMoney(string n, double m, double r, int t)
	{
		setAll(n, m, r, t);
	}
	myMoney()
	{
		setAll("", 0, 0, 0);
	}

};
class BigMoney : public myMoney
{
public:
	int periods;
	double getMoney()
	{
		double s = money;
		for (int k = 1;k <= time*periods;k++)
		{
			s *= (1 + rate / 100 / periods);
		}
		return s;
	}
	void showAll()
	{
		cout << "���: " << name << endl;
		cout << "�����: " << money << endl;
		cout << "������ (%): " << rate << endl;
		cout << "������ (���): " << time << endl;
		cout << "���������� � ���: " << periods<< endl;
		cout << "�������� �����: " << getMoney() << endl;
		for (int k = 1;k <= 35;k++)
		{
			cout << "-";
		}
		cout << endl;
	}
	void setAll(string n, double m, double r, int t, int p)
	{
		myMoney::setAll(n, m, r, t);
		periods = p;
	}
	BigMoney(string n, double m, double r, int t, int p = 1) : myMoney(n, m, r, t)
	{
		periods = p;
	}
	BigMoney(): myMoney() 
	{
	periods = 1;
	}
};


int main()
{
	setlocale(LC_ALL, "Russian");
	myMoney objA("��� ���������", 1200, 8, 5);
	BigMoney objB("���� Ը���", 1000, 7, 6, 2);
	BigMoney objC("��� �����", 1500, 6, 8);
	BigMoney objD;
	objD.setAll("��������� ������", 800, 10, 3, 4);
	objA.showAll();
	cout << endl;
	objB.showAll();
	cout << endl;
	objC.showAll();
	cout << endl;
	objD.showAll();
	cout << endl;



	system("pause>nul");
	return 0;
}

