------------------------------------------------------------------------------------------------------------------------
-- Data de criação: 06/05/2019
-- Criado por: Felipe Venancio de Sena
-- Descrição: Escreve o conteúdo de uma tabela em um arquivo csv em um local determinado
------------------------------------------------------------------------------------------------------------------------
CREATE PROCEDURE STP_ESCREVE_TABELA_CSV
	(
		@BANCO VARCHAR(100), -- BANCO DE DADOS ONDE ESTÁ LOCALIZADO A TABELA (NECESSÁRIO ESTAR DENTRO DO MESMO SERVIDOR
		@TABELA VARCHAR(100), -- TABELA ONDE ESTÁ LOCALIZADO O CONTEUDO QUE SERÁ ESCRITO NO ARQUIVO CSV
		@DESTINO VARCHAR(MAX) -- CAMINHO COMPLETO DO ARQUIVO QUE SERÁ SALVO
	)
AS
BEGIN
	SET NOCOUNT ON

	------------------------------------------------------------------------------------------------------------------------
	-- REMOVENDO AS TABELAS TEMPORÁRIAS
	------------------------------------------------------------------------------------------------------------------------
	IF OBJECT_ID('AREA_BACKUP..COLUNAS_TXT') IS NOT NULL DROP TABLE AREA_BACKUP..COLUNAS_TXT
	IF OBJECT_ID('AREA_BACKUP..CONTEUDO_TXT') IS NOT NULL DROP TABLE AREA_BACKUP..CONTEUDO_TXT
	IF OBJECT_ID('TEMPDB..#AUX') IS NOT NULL DROP TABLE #AUX

	------------------------------------------------------------------------------------------------------------------------
	-- CRIANDO A TABELA QUE RECEBERÁ O CONTEUDO DO ARQUIVO CSV
	------------------------------------------------------------------------------------------------------------------------
	CREATE TABLE AREA_BACKUP..CONTEUDO_TXT
		(
			TEXTO VARCHAR(MAX)
		)

	------------------------------------------------------------------------------------------------------------------------
	-- VARIÁVEIS DE TESTE
	------------------------------------------------------------------------------------------------------------------------
	--DECLARE @BANCO VARCHAR(100)
	--DECLARE @TABELA VARCHAR(100)
	--DECLARE @DESTINO VARCHAR(MAX)
	--SET @BANCO = 'AREA_BACKUP'
	--SET @TABELA = 'RELATORIO_WTTX_CIDADES'
	--SET @DESTINO = '\\10.206.104.223\C$\IMPORT\RelatorioCidades.csv'

	------------------------------------------------------------------------------------------------------------------------
	-- VARIÁVEIS INTERNAS
	------------------------------------------------------------------------------------------------------------------------
	DECLARE @TABELA_BANCO VARCHAR(202)
	SET @TABELA_BANCO = @BANCO+'.DBO.'+@TABELA

	------------------------------------------------------------------------------------------------------------------------
	-- SELECIONANDO O NOME DAS COLUNAS QUE SERÃO EXTRAÍDAS
	------------------------------------------------------------------------------------------------------------------------
	DECLARE @CONSULTA_COLUNAS VARCHAR(MAX)

	SET @CONSULTA_COLUNAS =
	'SELECT
		NAME,
		ROW_NUMBER() OVER(ORDER BY OBJECT_ID ASC) AS ID
	INTO
		AREA_BACKUP..COLUNAS_TXT
	FROM
		'+@BANCO+'.SYS.COLUMNS WITH (NOLOCK)
	WHERE
		OBJECT_ID = OBJECT_ID('''+@TABELA_BANCO+''')'

	PRINT @CONSULTA_COLUNAS
	EXEC (@CONSULTA_COLUNAS)
	------------------------------------------------------------------------------------------------------------------------
	-- PREENCHENDO OS TÍTULOS DO RELATÓRIO
	------------------------------------------------------------------------------------------------------------------------
	DECLARE @CONT INT
	SELECT
		@CONT = COUNT(0)
	FROM
		AREA_BACKUP..COLUNAS_TXT WITH (NOLOCK)

	SELECT
		*
	INTO
		#AUX
	FROM
		AREA_BACKUP..COLUNAS_TXT WITH (NOLOCK)

	DECLARE @TITULOS VARCHAR(MAX)
	DECLARE @ID INT

	WHILE @CONT > 0
		BEGIN
			SELECT TOP 1
				@TITULOS = ISNULL(@TITULOS,'')+NAME+';',
				@ID = ID
			FROM
				#AUX
			
			DELETE
				#AUX
			FROM
				#AUX
			WHERE
				ID = @ID
			
			SET @CONT = @CONT - 1
		END

	INSERT INTO AREA_BACKUP..CONTEUDO_TXT VALUES (@TITULOS)

	------------------------------------------------------------------------------------------------------------------------
	-- PREENCHENDO O CONTEÚDO DO RELATÓRIO
	------------------------------------------------------------------------------------------------------------------------
	DECLARE @CONSULTA VARCHAR(MAX)
	DECLARE @COLUNAS_CONSULTA VARCHAR(MAX)
	SET @CONSULTA = 'SELECT '

	SELECT
		@CONT = COUNT(0)
	FROM
		AREA_BACKUP..COLUNAS_TXT WITH (NOLOCK)

	WHILE @CONT > 0
		BEGIN
			SELECT TOP 1
				@COLUNAS_CONSULTA = ISNULL(@COLUNAS_CONSULTA,'')+'ISNULL(REPLACE(CONVERT(VARCHAR(MAX),['+NAME+']),'';'','',''),'''')+'';''+',
				@ID = ID
			FROM
				AREA_BACKUP..COLUNAS_TXT
			
			DELETE
				AREA_BACKUP..COLUNAS_TXT
			FROM
				AREA_BACKUP..COLUNAS_TXT
			WHERE
				ID = @ID
			
			SET @CONT = @CONT - 1
		END
	SET @COLUNAS_CONSULTA = LTRIM(RTRIM(@COLUNAS_CONSULTA))

	IF RIGHT(@COLUNAS_CONSULTA,1) = ','
		BEGIN
			SET @COLUNAS_CONSULTA = SUBSTRING(@COLUNAS_CONSULTA, 0, LEN(@COLUNAS_CONSULTA))
		END

	IF RIGHT(@COLUNAS_CONSULTA,1) = '+'
		BEGIN
			SET @COLUNAS_CONSULTA = SUBSTRING(@COLUNAS_CONSULTA, 0, LEN(@COLUNAS_CONSULTA))
		END

	SET @CONSULTA = @CONSULTA + @COLUNAS_CONSULTA + ' FROM ' + @TABELA_BANCO + ' WITH (NOLOCK)'

	PRINT @CONSULTA
	INSERT INTO AREA_BACKUP..CONTEUDO_TXT
	EXEC (@CONSULTA)

	------------------------------------------------------------------------------------------------------------------------
	-- ESCREVENDO O ARQUIVO CSV
	------------------------------------------------------------------------------------------------------------------------
	DECLARE @TEXTO AS VARCHAR(8000)
	SET @TEXTO = ' BCP '
		+ ' " SELECT TEXTO FROM AREA_BACKUP..CONTEUDO_TXT " '
		+ ' QUERYOUT '
		+ ' ' + @DESTINO + ' '
		+ ' -c -C1252 '
		+ ' -UADM '
		+ ' -PADMUSR '

	EXEC MASTER.DBO.XP_CMDSHELL @TEXTO
	
	------------------------------------------------------------------------------------------------------------------------
	-- REMOVENDO AS TABELAS TEMPORÁRIAS
	------------------------------------------------------------------------------------------------------------------------
	IF OBJECT_ID('AREA_BACKUP..COLUNAS_TXT') IS NOT NULL DROP TABLE AREA_BACKUP..COLUNAS_TXT
	IF OBJECT_ID('AREA_BACKUP..CONTEUDO_TXT') IS NOT NULL DROP TABLE AREA_BACKUP..CONTEUDO_TXT
	IF OBJECT_ID('TEMPDB..#AUX') IS NOT NULL DROP TABLE #AUX

END