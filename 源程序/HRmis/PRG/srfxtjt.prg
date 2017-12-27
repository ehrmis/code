SELECT 单位,工资,奖金,合计,人均收入,人均增资,增幅,人均奖金;
           FROM 收入增幅;
           INTO CURSOR SYS(2015)
           DO (_GENGRAPH) WITH 'QUERY'     