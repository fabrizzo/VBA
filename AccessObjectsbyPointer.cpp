// MassiveAndPointer.cpp: ���������� ����� ����� ��� ����������� ����������.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <cstdio>
#include <string>

using namespace std;
class MyClass {
public:
	string name;
	int number;
	void show() {
		cout << "���� ���:" << name << endl;
		cout << "���� �����:" << number << endl;
		for (int k = 0;k < 35;k++)
		{
			cout << "-";
		}
		cout << "\n";
	}
};

int main()
{
	setlocale(LC_ALL, "Russian");
	MyClass objA, objB;
	MyClass* p;
	p = &objA;
	p->name = "�������";
	p->number = 111;
	p->show();
	p = &objB;
	p->name = "�����";
	p->number = 222;
	p->show();
	cout << "��������� �������\n";
	objA.show();
	objB.show();
	system("pause>nul");
	return 0;
}

