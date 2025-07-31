WITH BOM_CTE AS (
    -- 锚点成员：获取直接子件
    SELECT 
        ID, 
        PARENT_ITEM_CODE, 
        ITEM_CODE, 
        QUANTITY,
        1 AS LEVEL  -- 层级，1表示直接子件
       ,CAST(bom.PARENT_ITEM_CODE + ' -> ' + bom.ITEM_CODE AS NVARCHAR(MAX)) AS BOM_Path
    FROM BOM
    WHERE PARENT_ITEM_CODE = '9C63647'
    
    UNION ALL
    
    -- 递归成员：获取下一级子件
    SELECT 
        t.ID, 
        t.PARENT_ITEM_CODE, 
        t.ITEM_CODE, 
        t.QUANTITY,
        c.LEVEL + 1 AS LEVEL  -- 层级递增
        , CAST(c.BOM_Path + ' -> ' + t.ITEM_CODE AS NVARCHAR(MAX))
    FROM BOM t
    INNER JOIN BOM_CTE c ON t.PARENT_ITEM_CODE = c.ITEM_CODE
)
-- 查询所有层级的子件
SELECT 
    ID, 
    PARENT_ITEM_CODE, 
    ITEM_CODE, 
    QUANTITY,
    LEVEL,BOM_Path
FROM BOM_CTE
ORDER BY LEVEL, PARENT_ITEM_CODE, ITEM_CODE;

-- get all purchase order of BOM components

