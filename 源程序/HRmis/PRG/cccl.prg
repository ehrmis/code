***************************
* .\CCCL.PRG
***************************
PROCEDURE CCCL
parameters E_ERRO , E_MESA , E_MESS1 , E_SYS16 , E_LINE
if  not empty(E_ERRO)
   if E_ERRO<>125
 set shadows on
 define window temp FROM INT((SROWS()-15)/2)  , INT((SCOLS()-60)/2)  TO INT((SROWS()+15)/2) , INT((SCOLS()+60)/2) colo 4+/6 float shadow title '�� ѡ ��[���á���������������]:';
 footer '[����������]'
 activate WINDOW temp
    @ 1 , 10 say '�� ѡ �� �� �� �� �� �� ��' font '����',15
    store 1 to DH53
    @ 5 , 3 get DH53 size 1 , 5 , 2 function '*h ��  ��;������;��  ��;��  ��' font '����',15
    wait window '�����ļ�:'+allt(E_SYS16)+'������:'+allt(str(E_LINE))+':'+allt(E_MESS1)+'�����:'+allt(str(E_ERRO))+'��������:'+allt(E_MESA) 
    READ    
    do case
    case DH53=1
  clear windows all 
      clear
      return
    case DH53=2
      CLEAR EVENTS      
    case DH53=3
      clear window all
      display memo
      display status
      display structure
      dir *.*
      wait
      clear
    case DH53=4
     CLEAR EVENTS
     quit
    endcase
  endif
endif
