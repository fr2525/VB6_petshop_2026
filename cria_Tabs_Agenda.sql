'***************************************************************************************
'***************************   CRIA A TABELA DE ANIMAIS   ******************************
'***************************************************************************************
'                              
sql = "CREATE TABLE tab_animais  (ID integer NOT NULL
								, Id_Cli integer NOT NULL
								, Nome  character(50) NOT NULL
								, Tipo_ani Int not null
								, dt_nasc date
								, observacoes varchar(200)
								, cuidados_especiais varchar(100)
								, foto varchar(100)
								, dt_ult_visita date
								, operador character(10)
								, dt_Atualiza timestamp
								, primary key (ID) )"
Cnn.Execute sql
Cnn.CommitTrans
        
'cria GENERATOR
sql = "CREATE GENERATOR GEN_ANI_ID1 "
Cnn.Execute sql
Cnn.CommitTrans

sql = "SET GENERATOR GEN_ANI_ID1 TO 0"
Cnn.Execute sql
Cnn.CommitTrans

sql = " CREATE TRIGGER TAB_ANIMAIS_BI FOR TAB_ANIMAIS ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_ANI_ID1, 1); END; "
Cnn.Execute sql
Cnn.CommitTrans


'****************************************************************************************
'****************   CRIA A TABELA DE TIPOS DE ANIMAL (CÃO/GATO/COELHO)  *****************
'****************************************************************************************
'                              
sql = "CREATE TABLE tab_tipos_an (ID integer NOT NULL
								, Descricao character(50) NOT NULL
								, pedigree CHAR(1)
								, operador character(10)
								, dt_Atualiza timestamp
								, primary key (ID) )"
Cnn.Execute sql
Cnn.CommitTrans
        
'cria GENERATOR
sql = "CREATE GENERATOR GEN_TPA_ID1 "
Cnn.Execute sql
Cnn.CommitTrans

sql = "SET GENERATOR GEN_TPA_ID1 TO 0"
Cnn.Execute sql
Cnn.CommitTrans

sql = " CREATE TRIGGER TAB_TIPOS_AN_BI FOR TAB_TIPOS_AN ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_TPA_ID1, 1); END; "
Cnn.Execute sql
Cnn.CommitTrans

'****************************************************************************************
'*****************  CRIA A TABELA DE SERVICOS - BANHO/TOSA/VACINAS/ETC  *****************
'****************************************************************************************
'                              
sql = "CREATE TABLE tab_servicos (ID integer NOT NULL
								, Descricao character(50) NOT NULL
								, valor NUMERIC(12,2)
								, TEMPO_EST NUMERIC(12,2)
								, operador character(10)
								, dt_Atualiza timestamp
								, primary key (ID) )"
Cnn.Execute sql
Cnn.CommitTrans
        
'cria GENERATOR
sql = "CREATE GENERATOR GEN_SERV_ID1 "
Cnn.Execute sql
Cnn.CommitTrans

sql = "SET GENERATOR GEN_SERV_ID1 TO 0"
Cnn.Execute sql
Cnn.CommitTrans

sql = " CREATE TRIGGER TAB_SERVICOS_BI FOR TAB_SERVICOS ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_SERV_ID1, 1); END; "
Cnn.Execute sql
Cnn.CommitTrans

'****************************************************************************************
'********************    CRIA A TABELA DE ATENDIMENTOS    *******************************
'****************************************************************************************
'                              
sql = "CREATE TABLE tab_atendimentos (Dt_atend timestamp NOT NULL
									, IdAnimal integer NOT NULL
									, Tipo_Atend integer NOT NULL
									, valor NUMERIC(12,2)
									, tempo_atendi NUMERIC(3,2)
									, operador character(10)
									, dt_Atualiza timestamp
									, primary key (dt_atend) )"
Cnn.Execute sql
Cnn.CommitTrans
        
'Não tem auto incremento porque o campo chave é TIMESTAMP


'****************************************************************************************
'********************       CRIA A TABELA DE VACINAS      *******************************
'****************************************************************************************
'                              
sql = "CREATE TABLE tab_vacinas (ID integer NOT NULL 
								,IdAnimal integer NOT NULL
								,Dt_atend timestamp NOT NULL
								,Descricao VARCHAR(100) NOT NULL
								,valor NUMERIC(12,2)
								,DT_PROXIMA DATE
								,operador character(10)
								,dt_Atualiza timestamp
								,primary key (idAnimal, dt_atend)  )"
Cnn.Execute sql
Cnn.CommitTrans
        
'cria GENERATOR
sql = "CREATE GENERATOR GEN_TVAC_ID1 "
Cnn.Execute sql
Cnn.CommitTrans

sql = "SET GENERATOR GEN_TVAC_ID1 TO 0"
Cnn.Execute sql
Cnn.CommitTrans

sql = " CREATE TRIGGER TAB_VACINAS_BI FOR TAB_VACINAS ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_TVAC_ID1, 1); END; "
Cnn.Execute sql
Cnn.CommitTrans


'****************************************************************************************
'********************       CRIA A TABELA DE PROMOCOES      *******************************
'****************************************************************************************
'                              
sql = "CREATE TABLE tab_promocoes (ID integer NOT NULL 
								  ,Dt_inicio timestamp  NOT NULL
								  ,Dt_fim timestamp NOT NULL
								  ,IdAnimal integer
								  ,IdTipoAten integer
								  ,Descricao VARCHAR(100) NOT NULL
								  ,Valor NUMERIC(12,2)
								  ,porcent NUMERIC(2,2)
								  ,operador character(10)
								  ,Dt_Atualiza timestamp
								  ,primary key (ID)  )
/*
Cnn.Execute sql
Cnn.CommitTrans

'cria GENERATOR
sql = "CREATE GENERATOR GEN_TPRO_ID1 "
Cnn.Execute sql
Cnn.CommitTrans

sql = "SET GENERATOR GEN_TPRO_ID1 TO 0"
Cnn.Execute sql
Cnn.CommitTrans

sql = " CREATE TRIGGER TAB_PROMOCOES_BI FOR TAB_PROMOCOES ACTIVE BEFORE INSERT POSITION 0 AS BEGIN "
sql = sql & "if (NEW.ID is NULL) then NEW.ID = GEN_ID(GEN_TPRO_ID1, 1); END; "
Cnn.Execute sql
Cnn.CommitTrans
*/



