CREATE PROCEDURE [dbo].[ProductDeliveries_Update_&_Transform]
(
    @ProductDeliveriesInfo NVARCHAR(128),           -- source database
    @ProductDeliveries NVARCHAR(128),                -- source table
    @ProdDel_Transformed_Tbls NVARCHAR(128),  -- target database
    @ProdDel_Bronze NVARCHAR(128)         -- target table (to be created)
)
AS
BEGIN
    DECLARE @SqlStatement NVARCHAR(MAX);

    -- Check if the source table exists in the source database
    IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @ProductDeliveries AND TABLE_SCHEMA = 'dbo')
    BEGIN
        -- Drop the target table if it exists in the target database
        IF OBJECT_ID(@ProdDel_Transformed_Tbls + '.dbo.' + @ProdDel_Bronze, 'U') IS NOT NULL
        BEGIN
            SET @SqlStatement =
                N'USE ' + QUOTENAME(@ProdDel_Transformed_Tbls) + N'; ' +
                N'DROP TABLE ' + QUOTENAME(@ProdDel_Bronze);

            -- Execute the dynamic SQL statement to drop the target table
            EXEC sp_executesql @SqlStatement;

            PRINT 'Target table dropped.';
        END

        -- Recreate the target table using SELECT INTO
        SET @SqlStatement =
            N'USE ' + QUOTENAME(@ProdDel_Transformed_Tbls) + N'; ' +
            N'SELECT * ' +
            N'INTO ' + QUOTENAME(@ProdDel_Bronze) + N' FROM ' + QUOTENAME(@ProductDeliveriesInfo) + N'.dbo.' + QUOTENAME(@VL06G);

        -- Execute the dynamic SQL statement to create the table and copy data
        EXEC sp_executesql @SqlStatement;

        PRINT 'Target table created with data from the source table.';
        
        -- Create ProdDel_TSA view in a separate dynamic SQL statement
        SET @SqlStatement = 
            N'USE ' + QUOTENAME(@ProdDel_Transformed_Tbls) +
            N'; EXEC(''CREATE VIEW dbo.[ProdDel_ALL] AS SELECT * FROM ' + QUOTENAME(@ProdDel_Bronze) + ''')';

        EXEC sp_executesql @SqlStatement;

        PRINT 'ProdDel_TSA view has been created successfully.';

        PRINT 'All the creation steps completed.';
    END
    ELSE
    BEGIN
        PRINT 'Source table does not exist in the source database.';
    END
END;
