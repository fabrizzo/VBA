// Generalized functions.cpp: ���������� ����� ����� ��� ����������� ����������.
//

#include "stdafx.h"
#include <iostream>
#include <cstdlib>
using namespace std;

template<class Z, class Y> void show(Z z, Y y) {
	cout << "������� � 2 �����������\n";
	cout << "�������� ��������� 1: " << z << endl;
	cout << "�������� ��������� 2: " << y << endl;
}


template<class X>void show(X x) {
	cout << "������� � ����� ����������\n";
	cout << "�������� ���������: " << x << endl;
}

int main()
{
	setlocale(LC_ALL, "Russian");
	show('A');
	show(123);
	show("�����");
	show(321, "�����");
	show('B', 456);
	show('C', 'D');

	



	system("pause>nul");
	return 0;
}
