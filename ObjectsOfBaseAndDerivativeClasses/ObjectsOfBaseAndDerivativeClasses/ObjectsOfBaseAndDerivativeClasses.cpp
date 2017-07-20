// ObjectsOfBaseAndDerivativeClasses.cpp: ���������� ����� ����� ��� ����������� ����������.
//

#include "stdafx.h"
#include <iostream>
#include <cstdlib>
using namespace std;
class Alpha {
public:
	char codeA;
	Alpha(char a) {
		codeA = a;
	}
	virtual void show() {
		cout << "����� �� ������ Alpha:" << codeA << endl;
	}
};
class Bravo {
public:
	char codeB;
	Bravo(char b) {
		codeB = b;
	}
	virtual void show() {
		cout << "����� �� ������ Bravo: " << codeB << endl;
	}
};
class Charlie : public Alpha, public Bravo {
public:
	Charlie(char a, char b) :Alpha(a), Bravo(b){}
	void show() {
		cout << "����� �� ������ Charlie: ";
		cout << codeA << codeB << endl;
	}
};






int main()
{
	setlocale(LC_ALL, "Russian");
	Alpha objA('A');
	objA.show();
	Bravo objB('B');
	objB.show();
	Charlie objC('C', 'D');
	objC.show();
	objA = objC;
	objB = objC;
	objA.show();
	objB.show();
	cout << "���������� ���������\n";
	Alpha* pntA = &objC;
	Bravo* pntB = &objC;
	Charlie* pntC = &objC;
	pntA->show();
	pntB->show();
	pntC->show();
	system("pause>nul");
    return 0;
}

