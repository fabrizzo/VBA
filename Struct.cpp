// MassiveAndPointer.cpp: определяет точку входа для консольного приложения.
//
#include "stdafx.h"
#include <iostream>
#include <cstdlib>
#include <cstdio>
#include <string>

using namespace std;

struct Book
{
	int id;
	int pages;
	string author;
	float cost;
};


int main()
{
	setlocale(LC_ALL, "Russian");
	
	Book b1;
	b1.id = 111;
	b1.pages = 300;
	b1.author = "Artur Connan Doil";
	b1.cost = 350.30;
	cout << "Id: " << b1.id << endl;
	cout << "Pages: " << b1.pages << endl;
	cout << "Author: " << b1.author << endl;
	cout << "Cost: " << b1.cost << endl;
	system("pause>nul");
	return 0;
}

