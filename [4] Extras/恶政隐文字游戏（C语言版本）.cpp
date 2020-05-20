#include <iostream>
#include <string.h>
#include <conio.h>
#include <windows.h>
using namespace std;

int main()
{
	system ("mode con cols=50 lines=10");
	system ("color f0");
	system ("title 恶政隐文字游戏　C语言版本　v20200520");

	Reset:
		int totalcount=0, correctcount=0, combocount=0, mistakecount=0,
            randomnumber1=0, randomnumber2=0, questionnumber=0, optionchosen=0;
        char input;
        string worddata1[151+1] = {"!!", "续", "蛤", "青", "改", "夏", "吉", "谦", "另", "高", "苟", "赛", "吼", "基", "钦", "无", "奉", "滋", "削", "图", "身", "西", "华", "谈", "风", "姿", "识", "捉", "跑", "森", "上", "拿", "抱", "长", "经", "碰", "闷", "发", "坠", "负", "特", "连", "要", "表", "民", "新", "批", "安", "不", "得", "い", "祈", "小", "维", "猪", "庆", "包", "吸", "禁", "倒", "星", "轻", "易", "通", "宽", "金", "律", "绿", "颐", "气", "冰", "岿", "大", "掀", "池", "风", "雨", "萨", "格", "尔", "吃", "麦", "十", "山", "二", "百", "换", "突", "满", "喷", "梁", "沼", "精", "细", "工", "八", "撸", "自", "困", "艰", "奋", "苦", "逆", "没", "发", "时", "读", "书", "闹", "清", "应", "神", "敬", "坡", "汹", "找", "瞻", "游", "亲", "谭", "麻", "泼", "膜", "品", "赵", "共", "称", "言", "粉", "五", "网", "干", "反", "翻", "一", "胡", "法", "坏", "逼", "支", "抖", "辣", "厉", "墙", "六", "坦", "铁", "螳", "当", "腊", "耿", "战"},
               worddata2[151+1] = {"!!", "命", "蟆", "蛙", "变", "威", "他", "虚", "请", "明", "利", "艇", "啊", "本", "点", "可", "告", "磁", "习", "样", "经", "方", "莱", "笑", "生", "势", "得", "急", "快", "破", "台", "衣", "歉", "者", "验", "到", "声", "财", "吼", "泽", "首", "任", "要", "态", "白", "闻", "判", "轨", "行", "罪", "よ", "翠", "熊", "尼", "头", "丰", "子", "精", "评", "车", "瀚", "关", "道", "商", "衣", "科", "玉", "玉", "使", "指", "棒", "然", "海", "翻", "塘", "狂", "骤", "格", "尔", "王", "饱", "子", "里", "路", "百", "斤", "肩", "开", "脸", "粪", "家", "气", "甚", "腻", "笔", "千", "袖", "息", "难", "苦", "斗", "吃", "差", "有", "酵", "代", "过", "单", "欢", "单", "验", "明", "畏", "涛", "涌", "准", "养", "泳", "自", "德", "批", "鸡", "乎", "韭", "弹", "惨", "帝", "论", "蛆", "毛", "紧", "五", "贼", "车", "派", "言", "轮", "球", "站", "乎", "阴", "椒", "害", "国", "四", "克", "骑", "臂", "车", "肉", "爽", "狼"},
               questionword, answerword, optionword[3+1];

	Title:
		system ("cls");
		cout<< "按 7 开始游戏，按 9 退出\n\n";

		input = getch();
		if      (input=='7') {goto Play;}
		else if (input=='9') {goto Exit;};
		goto Title;

    Play:
        //指定题面...
        randomnumber1 = (rand()%151)+1; questionnumber = randomnumber1;
        questionword = worddata1[randomnumber1]; answerword = worddata2[randomnumber1];

        //指定选项...
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
		cout<< "按 1 2 3 来选择答案，按 9 结束游戏\n\n";
        cout<< "总计数"<<totalcount<< "　正解数"<<correctcount<< "　"<< combocount<<"连击　失误数"<<mistakecount<< "\n\n";
        cout<< "题面：   "<<questionword<< "     （抽选中第"<<questionnumber<<"/151个字）\n\n";
        cout<< "选项：   1 "<<optionword[1]<< "   2 "<<optionword[2]<< "   3 "<<optionword[3]<< "\n\n";

		input = getch();
		if      (input=='9') {goto Reset;}
		else if (input=='1') {optionchosen=1; goto Judge;}
		else if (input=='2') {optionchosen=2; goto Judge;}
		else if (input=='3') {optionchosen=3; goto Judge;};
		goto Display;

    Judge:
        if (optionword[optionchosen]==answerword) {totalcount++; correctcount++; combocount++;}
            else {totalcount++; mistakecount++; combocount=0;
                  cout<< "错误！正确答案是「"<<answerword<<"」。请按任意键继续…\n"; getch();};
        goto Play;

    Exit:
        return 0;
};
