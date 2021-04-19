USE [MyDatabase]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

--====================================================================================================
-- Author: Camilo Sáez
-- Creation Date    : 18-04-2021
-- Description    : pr_DynamicExcel
--
-- ====================================================================================================

CREATE PROCEDURE [tabla_intermedia].[pr_DynamicExcel] @SHEET NVARCHAR(MAX),
                                                   @EXCEL NVARCHAR(MAX)
AS
    BEGIN
        SET NOCOUNT ON;
        BEGIN TRANSACTION;
        BEGIN TRY
		
            IF OBJECT_ID('MyDatabase.MySchema.DynamicExcel ', 'U') IS NOT NULL
                DROP TABLE MyDatabase.MySchema.DynamicExcel;
            DECLARE @SQLStr VARCHAR(MAX)= '
											SELECT *
											INTO MyDatabase.MySchema.DynamicExcel
											FROM OPENROWSET(''Microsoft.ACE.OLEDB.12.0'', ''Excel 12.0 Xml;Database=' + @EXCEL + ';'', ' + @SHEET + ') E';
            EXECUTE (@SQLStr);
            IF @@TRANCOUNT > 0
                COMMIT TRANSACTION;
				
        END TRY
        BEGIN CATCH
            ROLLBACK TRANSACTION;
            SELECT @@ERROR AS OUT_RETURN,
                   ERROR_NUMBER() AS OUT_NUMBER,
                   ERROR_MESSAGE() AS OUT_MESSAGE,
                   ERROR_SEVERITY() AS OUT_SEVERITY,
                   ERROR_STATE() AS OUT_STATE,
                   ERROR_PROCEDURE() AS OUT_PROCEDURE,
                   ERROR_LINE() AS OUT_LINE;
        END CATCH;
    END;