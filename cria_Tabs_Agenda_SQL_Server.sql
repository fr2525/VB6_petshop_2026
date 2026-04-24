--  tab_clientes  definição

-- Drop table

-- DROP TABLE  tab_clientes ;

CREATE TABLE  tab_clientes  (

	 id INTEGER AUTO_INCREMENT primary key, 
	NOME VARCHAR(250) NOT NULL,
	ENDERECO VARCHAR(250),
	BAIRRO VARCHAR(250),
	CIDADE VARCHAR(250),
	ESTADO VARCHAR(2),
	CEP VARCHAR(9),
	CGCCPF VARCHAR(16),
	RG VARCHAR(10),
	TELEFONE VARCHAR(16),
	CELULAR VARCHAR(16),
	ULTCOMPRA date,
	OBSERVA VARCHAR(250),
	 email  VARCHAR(250),
	 Operador  VARCHAR(10),
	 Datatual  DATE,
	 Contato  VARCHAR(50),
	 Limite  DECIMAL(18,2),
	 Saldo  DECIMAL(18,2),
	 EndCobra  VARCHAR(250),
	 BairCobra  VARCHAR(250),
	 CidaCobra  VARCHAR(50),
	 UFcobra  VARCHAR(2),
	 CepCobra  VARCHAR(9),
	 Negativo  BOOLEAN,
	 Insc_est  VARCHAR(20),
	DIAANIVER VARCHAR(2),
	MESANIVER VARCHAR(2),
	ANOANIVER VARCHAR(4)
);


/* '***************************************************************************************
'***************************   CRIA A TABELA DE ANIMAIS   ******************************
'***************************************************************************************
'  
*/                            
CREATE TABLE tab_animais  (ID integer identity NOT NULL
						,Id_Cli integer PRIMARY KEY AUTO_INCREMENT
						,Nome  character(50) NOT NULL
						,Tipo_ani Int not null
						,dt_nasc date
						,pedigree CHAR(1)						
						,observacoes varchar(200)
						,cuidados_especiais varchar(100)
						,foto varchar(100)
						,dt_ult_visita date
						,operador character(10)
						,dt_Atualiza datetime )
/*
'****************************************************************************************
'****************   CRIA A TABELA DE TIPOS DE ANIMAL (CÃO/GATO/COELHO)  *****************
'****************************************************************************************
' 
*/                             
CREATE TABLE tab_tipos_an (ID integer PRIMARY KEY AUTO_INCREMENT NOT NULL
						,Descricao varchar(50) NOT NULL
						,operador char(10)
						,dt_Atualiza datetime
						 )

/*
'****************************************************************************************
'*****************  CRIA A TABELA DE SERVICOS - BANHO/TOSA/VACINA/ETC   *****************
'****************************************************************************************
' */                              

CREATE TABLE tab_servicos (ID integer AUTO_INCREMENT PRIMARY KEY NOT NULL
					    ,Descricao character(50) NOT NULL
						,valor NUMERIC(12,2)
						,TEMPO_EST NUMERIC(12,2)
						,operador character(10)
						,dt_Atualiza datetime
						 )
/*
'****************************************************************************************
'********************    CRIA A TABELA DE ATENDIMENTOS    *******************************
'****************************************************************************************
' 
*/
                             
CREATE TABLE tab_atendimentos (Dt_atend datetime KEY NOT NULL
							,IdAnimal integer NOT NULL
							,Tipo_Atend integer NOT NULL
							,valor NUMERIC(12,2), tempo_atend NUMERIC(3,2)
							,operador character(10)
							,dt_Atualiza datetime
							 )
        
/*
'Não tem auto incremento porque o campo chave é TIMESTAMP

'****************************************************************************************
'********************       CRIA A TABELA DE VACINAS      *******************************
'****************************************************************************************
' 
*/                             
CREATE TABLE tab_vacinas (ID integer PRIMARY KEY AUTO_INCREMENT not null 
						,IdAnimal integer NOT NULL
						,Dt_atend datetime NOT NULL
						,Descricao VARCHAR(100) NOT NULL
						,Valor NUMERIC(12,2)
						,DT_PROXIMA DATE
						,operador character(10)
						,dt_Atualiza datetime
						,primary key (ID)  )


CREATE TABLE tab_promocoes (ID integer PRIMARY KEY AUTO_INCREMENT NOT NULL 
							,Dt_inicio datetime NOT NULL
							,Dt_fim datetime NOT NULL
							,IdAnimal integer
							,IdTipoAten integer
							,Descricao VARCHAR(100) NOT NULL
							,Valor NUMERIC(12,2)
							,porcent NUMERIC(2,2)
							,operador character(10)
							,Dt_Atualiza datetime
							,primary key (ID)  )

