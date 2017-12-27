DO CASE
  CASE ERROR()=1
  XX=MESSAGEBOX("输入/输出错误，您的磁盘没有放好或磁盘已满",0+48,"信息")
  人可='失败'
  CASE ERROR()=1002
  XX=MESSAGEBOX("输入/输出错误，您的磁盘没有放好或磁盘已满",0+48,"信息")
     RETURN 
     人可='失败'
  CASE ERROR()=56
  XX=MESSAGEBOX("没有足够的磁盘空间，操作不成功",0+48,"信息") 
  人可='失败'
 OTHER
  do CCCL with error() , message() , message(1) , sys(16) , lineno()
ENDCASE
retu