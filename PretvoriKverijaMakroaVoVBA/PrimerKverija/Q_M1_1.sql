SELECT Q_M1.OBRAZEC AS Expr1, Q_M1.GOD AS Expr2, Q_M1.RBR AS Expr3, [p1]-[p2] AS kontrola1, [p1]-[p3] AS kontrola2, [p1]-[p4] AS kontrola3, Q_M1.P1 AS Expr4, Q_M1.P2 AS Expr5, Q_M1.P3 AS Expr6, [Q_M1].P4 AS Expr1
FROM Q_M1
WHERE ((([p1]-[p2])<>0)) OR ((([p1]-[p3])<>0)) OR ((([p1]-[p4])<>0)) OR ((([Q_M1].[P1]) Is Null)) OR ((([Q_M1].[P2]) Is Null)) OR ((([Q_M1].[P3]) Is Null)) OR ((([Q_M1].[P4]) Is Null));

