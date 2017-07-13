// MassiveAndPointer.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <cstdio>
#include <string>

using namespace std;
class MyClass {
public:
	char code;
	MyClass* next;
	~MyClass() {
		cout << "Объект с полем " << code << " удален\n";
	}
	void show() {
		cout << code << " ";
		if (next) {
			next->show();
		}
	}
};

void deleteAll(MyClass* q) {
if (q->next) {
	deleteAll(q->next);
}
	delete q;
}

int main()
{
	setlocale(LC_ALL, "Russian");
	int n = 25;
	MyClass* pnt = new MyClass;
	pnt->code = 'A';
	MyClass *p;
	p = pnt;
	for (int k = 1; k <= n;k++) {
		p->next = new MyClass;
		p->next->code = p->code + 1;
		p = p->next;
	}
	p->next = 0;
	pnt->show();
	cout << endl;
	deleteAll(pnt);
	
	system("pause>nul");
	return 0;
}

