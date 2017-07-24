// Generalizer Class.cpp: ���������� ����� ����� ��� ����������� ����������.
//

#include "stdafx.h"
#include <iostream>
#include <string>
using namespace std;
template<class A, class B> class MyClass {
public: 
	A first;
	B second;
	MyClass(A f, B s) {
		first = f;
		second = s;
	}
	void show() {
		cout << "������ ���� " << first << endl;
		cout << "������ ���� " << second << endl;
	}
};
int main()
{
	setlocale(LC_ALL, "Russian");
	MyClass<int, char> objA(100, 'A');
	MyClass<string, double> objB("�����", 10.5);
	MyClass<char*, string> objC("������", "������");
	MyClass<int, int> objD(1, 2);
	objA.show();
	objB.show();
	objC.show();
	system("pause>nul");
    return 0;
}

