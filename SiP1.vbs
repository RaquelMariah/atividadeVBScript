dim o 'variavel para sorteio operadores
dim num1 'variavel para sorteio dos numeros
dim num2 'variavel para sorteio dos numeros
dim cont 'variavel para contador laco repeticao numeros
dim i		'variavel para contador laco repeticao operadores
dim sorteioNumero(10) 'vetor para 10 numeros
dim sorteioOperador(3) 'vetor para tres opera��es
dim nome


call sorteio1() '-------------------SORTEIO DOS NUMEROS

'call input1()

sub input1()
scase=cstr(inputbox("=================================" + vbNewLine & _
"ACERTE O CALCULO MATEM�TICO" + vbNewLine & _
"=================================" + vbNewLine & _
"	RESOLVA: 10 * 9 = ??? " + vbNewLine & _ 
"=================================", "SISTEMAS DE INFORMA��O - P1") + vbokcancel)
end sub

sub sorteio1() '-------------------SORTEIO DOS NUMEROS 1
'----LOG
msgbox("sorteio dos numeros (1)")
cont=1 
for cont = 1 to 10 step 0 'para diminuir acrescentar - antes do numero final'
	randomize(second(time)) 'padr�o'
	num1=int(rnd * 11) +1
	cont = cont + 1
	msgbox(num1)
next
call sorteio2() '-------------------SORTEIO DOS NUMEROS 2
end sub


sub sorteio2() '-------------------SORTEIO DOS NUMEROS 2
'----LOG
msgbox("sorteio dos numeros (2)")
cont=1 
for cont = 1 to 10 step 0 'para diminuir acrescentar - antes do numero final'
	randomize(second(time)) 'padr�o'
	num2=int(rnd * 11) +1
	cont = cont + 1
	msgbox(num2)
next
call sorteio3() '-------------------SORTEIO DOS OPERADORES
end sub

sub sorteio3() '-------------------SORTEIO DOS OPERADORES
'----LOG
msgbox("sorteio dos operadores (3)")
i = 1
for i=1 to 3 step 0 
	sorteioOperador(1)="ADI��O"
	sorteioOperador(2)="SUBTRA��O"
	sorteioOperador(3)="MULTIPLICA��O"
	randomize(second(time))
	o=int(rnd*4)+1
	msgbox(sorteioOperador(o))
	i = i + 1
next	
end sub