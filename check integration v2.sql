WITH noNull AS (SELECT COALESCE(rmsk, '0') AS rmsk FROM Tst WHERE rmsk like '%RS'),

		 t1 AS (SELECT rmsk,
			LEFT(rmsk, 4) AS StoreID,
			SUBSTRING(RMSk, CHARINDEX('-', RMSk) + 1,
				CHARINDEX('-', RMSk, CHARINDEX('-', RMSk) + 1) - CHARINDEX	('-', RMSk) - 1) AS invno
			FROM noNull),
		t2 AS (SELECT DISTINCT rmsk, storeID , convert(int,invNO) AS InvNO FROM t1),
		t3 AS (SELECT *, ROW_NUMBER() OVER(partition by storeid ORDER BY invNO) RowNo FROM t2),
		t4 as (	SELECT
					storeID,
					min(invNO) AS the_1st_Inv,
					max(invNO) AS last_Inv,	
					count(RowNo) AS exist_INV,
					max(invNO) - min(invNO) + 1 AS diff
				FROM t3
				GROUP BY storeID)

select *,
		case when exist_INV = diff then ' Successfull' else 'Error' end as integration_status
from t4
order by 6
