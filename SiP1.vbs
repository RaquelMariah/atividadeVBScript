dim o 'variavel para sorteio operadores
dim num1 'variavel para sorteio dos numeros
dim num2 'variavel para sorteio dos numeros
dim cont 'variavel para contador laco repeticao numeros
dim i		'variavel para contador laco repeticao operadores
dim sorteioNumero(10) 'vetor para 10 numeros
dim sorteioOperador(3) 'vetor para tres opera��es
dim resultado 			'variavel para resultado das operacoes
dim resposta 			'variavel para armazenar resultado da operaco e exibir no front

dim simboloOperacao		'variavel para armazenar o simbolo da operacao e exibir no front
dim scase
dim acertos				'variavel para armazenar numero de acertos do usu�rio
dim resp 

call sorteio1() '-------------------SORTEIO DOS NUMEROS



sub sorteio1() '-------------------SORTEIO DOS NUMEROS 1
'----LOG
'msgbox("sorteio dos numeros (1)")
cont=1 
for cont = 1 to 10 step 0 'para diminuir acrescentar - antes do numero final'
	randomize(second(time)) 'padr�o'
	num1=int(rnd * 11) +1
	cont = cont + 1
	'msgbox(num1)
next
call sorteio2() '-------------------SORTEIO DOS NUMEROS 2
end sub


sub sorteio2() '-------------------SORTEIO DOS NUMEROS 2
'----LOG
'msgbox("sorteio dos numeros (2)")
cont=1 
for cont = 1 to 10 step 0 'para diminuir acrescentar - antes do numero final'
	randomize(second(time)) 'padr�o'
	num2=int(rnd * 11) +1
	cont = cont + 1
'	msgbox(num2)
next
call sorteio3() '-------------------SORTEIO DOS OPERADORES
end sub

sub sorteio3() '-------------------SORTEIO DOS OPERADORES
'----LOG
'msgbox("sorteio dos operadores (3)")
i = 1
for i=1 to 3 step 0 
	sorteioOperador(1)="ADI��O"
	sorteioOperador(2)="SUBTRA��O"
	sorteioOperador(3)="MULTIPLICA��O"
	randomize(second(time))
	o=int(rnd*3)+1
'	msgbox(sorteioOperador(o))
	i = i + 1
next	
if 	sorteioOperador(o) = sorteioOperador(1) then '----------------------------ADI��O
	call operacaoSoma()
	resposta = resultado
	
elseif sorteioOperador(o) = sorteioOperador(2) then '-------------------------SUBTRA��O
	call operacaoSubtracao()
	resposta = resultado
	
else 												 '-------------------------MULTIPLICA��O
	call operacaoMultiplicacao()
	resposta = resultado
	
end if
call input1
end sub
'-------------------------------------------------------OPERACOES
sub operacaoSoma()
	resultado = num1 + num2
	simboloOperacao = " + "

end sub

sub operacaoSubtracao()
	resultado = num1 - num2
	simboloOperacao = " - "

end sub

sub operacaoMultiplicacao()
	resultado = num1 * num2
	simboloOperacao = " * "

end sub
'-------------------------------------------------------OPERACOES

sub acertou()
	MsgBox("Parab�ns voc� acertou!" + vbNewLine & _
	"Qtde de acertos: " + ""& acertos &"")
	call sorteio1()
	
end sub

sub errou()

resp=msgbox("Deseja jogar novamente?", vbQuestion + vbYesNo, "ATEN��O")
	if resp = vbYes then
		call input1()
	else
		wscript.quit
	end if
	
end sub


sub input1()

scase=(cint(inputbox("=================================" + vbNewLine & _
"ACERTE O CALCULO MATEM�TICO" + vbNewLine & _
"=================================" + vbNewLine & _
"	RESOLVA: " + ""& num1 &"" + ""& simboloOperacao&"" +  ""& num2 &"" + " = " + vbNewLine & _ 
"=================================", "SISTEMAS DE INFORMA��O - P1")))
if scase = resposta then
	acertos = acertos + 1
	call acertou()
elseif scase = "" then
	msgbox("Op��o inv�lida!")
	call input1
else
	msgbox("Voc� errou!")
	call errou()
end if
end sub



