#include <iostream>
#include <string.h>
#include <conio.h>
#include <windows.h>
using namespace std;

int main()
{
	system ("mode con cols=50 lines=10");
	system ("color f0");
	system ("title ������������Ϸ��C���԰汾��v20200520");

	Reset:
		int totalcount=0, correctcount=0, combocount=0, mistakecount=0,
            randomnumber1=0, randomnumber2=0, questionnumber=0, optionchosen=0;
        char input;
        string worddata1[151+1] = {"!!", "��", "��", "��", "��", "��", "��", "ǫ", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "ͼ", "��", "��", "��", "̸", "��", "��", "ʶ", "׽", "��", "ɭ", "��", "��", "��", "��", "��", "��", "��", "��", "׹", "��", "��", "��", "Ҫ", "��", "��", "��", "��", "��", "��", "��", "��", "��", "С", "ά", "��", "��", "��", "��", "��", "��", "��", "��", "��", "ͨ", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "ʮ", "ɽ", "��", "��", "��", "ͻ", "��", "��", "��", "��", "��", "ϸ", "��", "��", "ߣ", "��", "��", "��", "��", "��", "��", "û", "��", "ʱ", "��", "��", "��", "��", "Ӧ", "��", "��", "��", "��", "��", "հ", "��", "��", "̷", "��", "��", "Ĥ", "Ʒ", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "һ", "��", "��", "��", "��", "֧", "��", "��", "��", "ǽ", "��", "̹", "��", "�", "��", "��", "��", "ս"},
               worddata2[151+1] = {"!!", "��", "�", "��", "��", "��", "��", "��", "��", "��", "��", "ͧ", "��", "��", "��", "��", "��", "��", "ϰ", "��", "��", "��", "��", "Ц", "��", "��", "��", "��", "��", "��", "̨", "��", "Ǹ", "��", "��", "��", "��", "��", "��", "��", "��", "��", "Ҫ", "̬", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "ͷ", "��", "��", "��", "��", "��", "�", "��", "��", "��", "��", "��", "��", "��", "ʹ", "ָ", "��", "Ȼ", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "·", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "ǧ", "��", "Ϣ", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "η", "��", "ӿ", "׼", "��", "Ӿ", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "ë", "��", "��", "��", "��", "��", "��", "��", "��", "վ", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "��", "ˬ", "��"},
               questionword, answerword, optionword[3+1];

	Title:
		system ("cls");
		cout<< "�� 7 ��ʼ��Ϸ���� 9 �˳�\n\n";

		input = getch();
		if      (input=='7') {goto Play;}
		else if (input=='9') {goto Exit;};
		goto Title;

    Play:
        //ָ������...
        randomnumber1 = (rand()%151)+1; questionnumber = randomnumber1;
        questionword = worddata1[randomnumber1]; answerword = worddata2[randomnumber1];

        //ָ��ѡ��...
        randomnumber2 = (rand()%3)+1;
        optionword[randomnumber2] = answerword;
        switch (randomnumber2)
        {
            case 1:
                do {randomnumber1 = (rand()%151)+1; optionword[2] = worddata2[randomnumber1];} while (optionword[2] == optionword[1]);
                do {randomnumber1 = (rand()%151)+1; optionword[3] = worddata2[randomnumber1];} while (optionword[3] == optionword[1] || optionword[3] == optionword[2]);
                break;
            case 2:
                do {randomnumber1 = (rand()%151)+1; optionword[1] = worddata2[randomnumber1];} while (optionword[1] == optionword[2]);
                do {randomnumber1 = (rand()%151)+1; optionword[3] = worddata2[randomnumber1];} while (optionword[3] == optionword[2] || optionword[3] == optionword[1]);
                break;
            case 3:
                do {randomnumber1 = (rand()%151)+1; optionword[1] = worddata2[randomnumber1];} while (optionword[1] == optionword[3]);
                do {randomnumber1 = (rand()%151)+1; optionword[2] = worddata2[randomnumber1];} while (optionword[2] == optionword[3] || optionword[2] == optionword[1]);
                break;
        };

    Display:
        system ("cls");
		cout<< "�� 1 2 3 ��ѡ��𰸣��� 9 ������Ϸ\n\n";
        cout<< "�ܼ���"<<totalcount<< "��������"<<correctcount<< "��"<< combocount<<"������ʧ����"<<mistakecount<< "\n\n";
        cout<< "���棺   "<<questionword<< "     ����ѡ�е�"<<questionnumber<<"/151���֣�\n\n";
        cout<< "ѡ�   1 "<<optionword[1]<< "   2 "<<optionword[2]<< "   3 "<<optionword[3]<< "\n\n";

		input = getch();
		if      (input=='9') {goto Reset;}
		else if (input=='1') {optionchosen=1; goto Judge;}
		else if (input=='2') {optionchosen=2; goto Judge;}
		else if (input=='3') {optionchosen=3; goto Judge;};
		goto Display;

    Judge:
        if (optionword[optionchosen]==answerword) {totalcount++; correctcount++; combocount++;}
            else {totalcount++; mistakecount++; combocount=0;
                  cout<< "������ȷ���ǡ�"<<answerword<<"�����밴�����������\n"; getch();};
        goto Play;

    Exit:
        return 0;
};
