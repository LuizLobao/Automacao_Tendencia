import segredos
from datetime import date, datetime
import pyodbc
import time



def ATIVAR_TEND_TABLEAU_teste_Jan22():
	comando_sql='''
                ALTER PROCEDURE [dbo].[SP_PC_CG_IND_Acompanhamento_Diario_Final] @ANOMES VARCHAR(6)
                WITH RECOMPILE
                AS
                --SET IMPLICIT_TRANSACTIONS ON
                BEGIN-- try 

                    INSERT INTO TBL_PC_TEMPO_PROCEDURES
                    SELECT
                    'SP_PC_CG_IND_Acompanhamento_Diario_Final' AS [PROCEDURE],
                    'INICIO' AS INI_FIM,
                    GETDATE() AS DATA_HORA
                    
                    
                    
                    
                    
                    PRINT '-----------------------------------------------------------------------------------------'
                    PRINT '-            INICIO CARGA SP_PC_CG_IND_Acompanhamento_Diario_Final   MÊS: '+@ANOMES
                    PRINT '-----------------------------------------------------------------------------------------'

                    DECLARE @ANOMES_M1 CHAR(6)
                    SET @ANOMES_M1 = dbo.format_date(DATEADD(MONTH, DATEDIFF(MONTH, 0 , GETDATE()-1)-1,0),'YYYYMM')


                    delete from TBL_IND_VAR_BASEMETA_PORDU where LEFT(data,6) = @ANOMES
                    /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                    --AND TIPO_INDICADOR <> 'TENDÊNCIA'

                    insert into TBL_IND_VAR_BASEMETA_PORDU
                    SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                CASE WHEN GRUPO_PLANO = 'OI GALERA PRÉ' THEN 'PRÉ-PAGO' ELSE GRUPO_PLANO END GRUPO_PLANO,
                                CASE WHEN ((COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                                ELSE C.CANAL_RB END CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.FILIAL,
                                'VAREJO' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BASERESULTADOS AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.FILIAL = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CANAISRELATORIOBERNARDO AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND GRUPO_PLANO NOT IN ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                        AND INDBD <> 'CANCELAMENTO' 
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                CASE WHEN GRUPO_PLANO = 'OI GALERA PRÉ' THEN 'PRÉ-PAGO' ELSE GRUPO_PLANO END,
                                DATA,
                                CASE WHEN ((COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                                ELSE C.CANAL_RB END,
                                R.REGIONAL,
                                B.FILIAL
                        UNION ALL
                        SELECT	
                                CASE WHEN (INDBD IN ('VL', 'MIGRACAO')) THEN 'VL' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.FILIAL,
                                'VAREJO' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BASERESULTADOS AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.FILIAL = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CANAISRELATORIOBERNARDO AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND  DATA = @ANOMES
                        AND GRUPO_PLANO IN ('CONECTADO', 'RESIDENCIAL', 'COMPLETO')
                        AND INDBD IN ('VL', 'MIGRACAO')
                        GROUP BY 
                                CASE WHEN (INDBD IN ('VL', 'MIGRACAO')) THEN 'VL' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.FILIAL
                        UNION ALL
                        SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.FILIAL,
                                'VAREJO' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BASERESULTADOS AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.FILIAL = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CANAISRELATORIOBERNARDO AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND GRUPO_PLANO IN ('CONECTADO', 'RESIDENCIAL', 'COMPLETO')
                        AND INDBD IN ('GROSS', 'MIGRACAO')
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.FILIAL
                    UNION ALL
                    SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.UF_CLIENTE FILIAL,
                                'EMP CLI' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BaseResultados_Empresarial AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.UF_CLIENTE = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND INDBD <> 'CANCELAMENTO' 
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.UF_CLIENTE
                        UNION ALL
                        SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.UF_CARTEIRA FILIAL,
                                'EMP PDV' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BaseResultados_Empresarial AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.UF_CARTEIRA = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */ 
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND INDBD <> 'CANCELAMENTO' 
                        AND REGIONAL_AGRUPADA IS NOT NULL
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.UF_CARTEIRA



                    DELETE FROM TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE WHERE DATA = @ANOMES
                    AND NOT (TIPO_INDICADOR = 'META' AND DATA IN ('202007', '202008', '202009') AND GRUPO_PLANO IN ('CONTROLE_BOLETO',
                    'CONTROLE_CARTAO', 'PÓS ALONE', 'PÓS OIT', 'PRÉ-PAGO', 'PRÉ-D3', 'MOVEL'))

                    insert into TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE
                    select 
                            TIPO_INDICADOR,
                            data,
                            INDBD,
                            GRUPO_PLANO,
                            CANAL, 
                            REGIONAL,
                            SEGMENTO,
                            CASE 
                                WHEN CANAL IN ('Smart Message','TLV Outros','Outros', 'Outros BRI','Condominios','','NULL','S2S','Outros EMP','S2S EMP','TLV Outros EMP' ) THEN 'Outros Nacionais' 
                                WHEN CANAL IN ('TLV Receptivo','TLV Ativo','TLV Ativo EMP','TLV Receptivo EMP' ) THEN 'TLV' 
                                WHEN CANAL IN ('WEB','WEB EMP') THEN 'WEB' 
                                WHEN CANAL = 'Anteneiros' THEN 'Anteneiros' 
                                WHEN CANAL IN ('TLV Receptivo BRI','TLV Outros BRI','TLV Ativo BRI','TLV BRI' ) THEN 'TLV BRI' 
                                WHEN CANAL LIKE '%EMP%' AND [SEGMENTO] <> 'VAREJO' THEN 'Gestão Regional EMP'
                                WHEN CANAL NOT LIKE '%EMP%' AND [SEGMENTO] <> 'VAREJO' THEN 'Gestão Regional VAR'
                            ELSE 'Gestão Regional' END GESTAO,
                            UF,
                            cast( DIA as varchar(2)) DIA,
                            META_DU QTD


                    from(	SELECT		distinct 
                            TIPO_INDICADOR,
                            a.data,
                            A.indbd INDBD,
                            CASE WHEN A.grupo_plano IN ('OI GALERA PRÉ', 'PRÉ-PAGO') THEN 'PRÉ-PAGO' ELSE A.grupo_plano END GRUPO_PLANO,
                            A.canal_rb CANAL, 
                            REG.REGIONAL_AGRUPADA REGIONAL,
                            SEGMENTO,
                            LTRIM(RTRIM(A.FILIAL)) UF,
                            '' AS DIA,
                            a.VALOR META_DU
                            -- ,A.VALOR
                            --,DU_MES_PRODUTO.VALOR DU_MES_PRODUTO
                            -- ,DU_MES.VALOR DU_MES
                            -- ,P.VALOR P
                        
                    FROM TBL_IND_VAR_BASEMETA_PORDU A
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS REG ON A.FILIAL = REG.FILIAL
                                    
                    WHERE --a.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                                a.grupo_plano not in (	'CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS',
                                                            'TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS',
                                                            'BANDA LARGA FIBRA C','BANDA LARGA FIBRA EO','BANDA LARGA FIBRA EX OBRA','BANDA LARGA FIBRA SI','FIBRA C',
                                                            'FIBRA EO','FIBRA EX OBRA','FIBRA SI','FIXO FIBRA C','FIXO FIBRA EO','FIXO FIBRA EX OBRA','FIXO FIBRA SI',
                                                            'OI TV FIBRA C','OI TV FIBRA EO','OI TV FIBRA EX OBRA','OI TV FIBRA SI','OIT COMERCIAL',
                                                            'OIT CONECTADO','OIT SOLUCAO COMPLETA','VADA')
                            AND A.TIPO_INDICADOR = 'TENDÊNCIA'
                            AND A.DATA = @ANOMES

                            UNION ALL


                    SELECT		distinct 
                            TIPO_INDICADOR,
                            a.data,
                            A.indbd INDBD,
                            CASE WHEN A.grupo_plano IN ('OI GALERA PRÉ', 'PRÉ-PAGO') THEN 'PRÉ-PAGO' ELSE A.grupo_plano END GRUPO_PLANO,
                            A.canal_rb CANAL, 
                            REG.REGIONAL_AGRUPADA REGIONAL,
                            SEGMENTO,
                            LTRIM(RTRIM(A.FILIAL)) UF,
                            RIGHT(C.ANOMESDIA,2) AS DIA,
                            case 
                                when A.VALOR = 0 then 0
                                when DU_MES.VALOR > 0 or DU_MES.VALOR is not null then (A.VALOR/DU_MES.VALOR)*C.VALOR
                                when DU_MES_PRODUTO.VALOR > 0 or DU_MES_PRODUTO.VALOR is not null then (A.VALOR/DU_MES_PRODUTO.VALOR)*p.VALOR 
                            else	0 end META_DU
                            -- ,A.VALOR
                            --,DU_MES_PRODUTO.VALOR DU_MES_PRODUTO
                            -- ,DU_MES.VALOR DU_MES
                            -- ,P.VALOR P
                        
                    FROM TBL_IND_VAR_BASEMETA_PORDU A
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS REG ON A.FILIAL = REG.FILIAL

                    /* TOTAL DU POR REGIONAL - CANAL */
                    left JOIN (
                                    SELECT 
                                        * 
                                    FROM TBL_IND_VAR_DU_SUMARIZADO
                                    where ANOMES = @ANOMES
                            ) DU_MES ON A.INDBD = DU_MES.INDBD AND 
                                                        --a.GRUPO_PLANO = DU_MES.PRODUTO_NOVO AND
                                                        case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = DU_MES.PRODUTO_NOVO AND
                                                        CASE WHEN A.CANAL_RB = 'Smart Message' THEN 'TLV Outros' 
                                                        ELSE A.CANAL_RB END = DU_MES.CANAL AND
                                                        A.REGIONAL = DU_MES.REGIONAL AND
                                                        A.DATA = DU_MES.ANOMES

                    /* DU POR DIA PARA REGIONAL - CANAL */
                    left JOIN (SELECT ANOMESDIA,
                                    left(ANOMESDIA,6) ANOMES,
                                    INDBD,
                                    CASE WHEN DU_DIA.PRODUTO IN ('OI GALERA PRÉ', 'PRÉ-PAGO') THEN 'PRÉ-PAGO' ELSE DU_DIA.PRODUTO END PRODUTO_NOVO,
                                    CANAL,
                                    REGIONAL,
                                    valor
                                FROM TBL_pc_du AS DU_DIA
                                            where left(ANOMESDIA,6) = @ANOMES
                                ) C ON 
                                        A.INDBD = C.INDBD AND 
                                        case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = C.PRODUTO_NOVO AND
                                        CASE WHEN A.CANAL_RB = 'Smart Message' THEN 'TLV Outros' ELSE
                                        A.CANAL_RB END = C.CANAL AND
                                        A.REGIONAL = C.REGIONAL and
                                        A.DATA = C.ANOMES
                    
                    /* TOTAL DU POR PRODUTOS */
                    left JOIN (
                                    SELECT 
                                            LEFT(ANOMESDIA,6) ANOMES,
                                            INDBD,
                                            PRODUTO_NOVO,
                                            SUM(cast(rtrim(ltrim(replace(VALOR,',','.'))) as float)) VALOR 
                                    FROM TBL_PC_DU_PRODUTO
                                    where left(ANOMESDIA,6) = @ANOMES
                                    group by 
                                            LEFT(ANOMESDIA,6),
                                            INDBD,
                                            PRODUTO_NOVO
                            ) DU_MES_PRODUTO ON	A.INDBD = DU_MES_PRODUTO.INDBD AND 
                                                    case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = DU_MES_PRODUTO.PRODUTO_NOVO AND
                                                    A.DATA = DU_MES_PRODUTO.ANOMES

                    /* DU POR DIA PARA PRODUTO */
                    left JOIN (SELECT 
                                        ANOMESDIA,
                                        LEFT(ANOMESDIA,6) ANOMES,
                                        INDBD,
                                        PRODUTO_NOVO,
                                        SUM(cast(rtrim(ltrim(replace(VALOR,',','.'))) as float)) VALOR 
                                FROM TBL_PC_DU_PRODUTO
                                where left(ANOMESDIA,6) = @ANOMES
                                group by 
                                        ANOMESDIA,
                                        LEFT(ANOMESDIA,6),
                                        INDBD,
                                        PRODUTO_NOVO
                                ) P ON 
                                        A.INDBD = P.INDBD AND 
                                        case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = P.PRODUTO_NOVO AND
                                        A.DATA = P.ANOMES
                    
                                    
                    WHERE --a.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                                a.grupo_plano not in (	'CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS',
                                                            'TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS',
                                                            'BANDA LARGA FIBRA C','BANDA LARGA FIBRA EO','BANDA LARGA FIBRA EX OBRA','BANDA LARGA FIBRA SI','FIBRA C',
                                                            'FIBRA EO','FIBRA EX OBRA','FIBRA SI','FIXO FIBRA C','FIXO FIBRA EO','FIXO FIBRA EX OBRA','FIXO FIBRA SI',
                                                            'OI TV FIBRA C','OI TV FIBRA EO','OI TV FIBRA EX OBRA','OI TV FIBRA SI','OIT COMERCIAL',
                                                            'OIT CONECTADO','OIT SOLUCAO COMPLETA','VADA')
                            AND A.TIPO_INDICADOR <> 'TENDÊNCIA'
                            AND A.DATA = @ANOMES
                            AND NOT (TIPO_INDICADOR = 'META' AND DATA IN ('202007', '202008', '202009') AND GRUPO_PLANO IN ('CONTROLE_BOLETO',
                            'CONTROLE_CARTAO', 'PÓS ALONE', 'PÓS OIT', 'PRÉ-PAGO', 'PRÉ-D3', 'MOVEL'))

                    union all

                    SELECT  'REAL' TIPO_INDICADOR,
                            left(data,6) anomes,
                            INDBD,	
                        GRUPO_PLANO,
                        CASE WHEN ((B.COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                        ELSE C.CANAL_RB END canal,
                        R.REGIONAL_AGRUPADA regional,
                        'VAREJO' SEGMENTO,
                        LTRIM(RTRIM(B.FILIAL)) AS UF,
                        RIGHT(DATA,2) AS DIA,
                        SUM(B.QTD) AS valor

                    FROM TBL_RE_BaseResultadoDiario B	
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS R ON B.FILIAL = R.FILIAL	
                    LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo AS C ON B.CANAL_FINAL = C.CANAL_FINAL	

                    WHERE --c.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                    b.grupo_plano not in ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                    AND B.INDBD <> 'CANCELAMENTO' 
                    AND LEFT(DATA,6) = @ANOMES
                    
                    GROUP BY INDBD,	
                        GRUPO_PLANO,
                        left(data,6),
                        CASE WHEN ((B.COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                        ELSE C.CANAL_RB END,
                        R.REGIONAL_AGRUPADA,
                    
                        LTRIM(RTRIM(B.FILIAL)),
                        RIGHT(DATA,2)

                    union all

                    SELECT  'REAL' TIPO_INDICADOR,
                            left(data,6) anomes,
                            INDBD,	
                        GRUPO_PLANO,
                        C.CANAL_RB canal,
                        R.REGIONAL_AGRUPADA regional,
                        'EMP CLI' SEGMENTO,
                        LTRIM(RTRIM(B.UF_CLIENTE)) AS UF,
                        RIGHT(DATA,2) AS DIA,
                        SUM(B.QTD) AS valor

                    FROM TBL_RE_BaseResultadoDiario_Empresarial B	
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS R ON B.UF_CLIENTE = R.FILIAL	
                    LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL	

                    WHERE --c.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                    b.grupo_plano not in ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                    AND B.INDBD <> 'CANCELAMENTO' 
                    AND LEFT(DATA,6)  = @ANOMES
                    
                    GROUP BY INDBD,	
                        GRUPO_PLANO,
                        left(data,6),
                        C.CANAL_RB,
                        R.REGIONAL_AGRUPADA,
                    
                        LTRIM(RTRIM(B.UF_CLIENTE)),
                        RIGHT(DATA,2)

                    union all

                    SELECT  'REAL' TIPO_INDICADOR,
                            left(data,6) anomes,
                            INDBD,	
                        GRUPO_PLANO,
                        C.CANAL_RB canal,
                        R.REGIONAL_AGRUPADA regional,
                        'EMP PDV' SEGMENTO,
                        LTRIM(RTRIM(B.UF_CARTEIRA)) AS UF,
                        RIGHT(DATA,2) AS DIA,
                        SUM(B.QTD) AS valor
                    FROM TBL_RE_BaseResultadoDiario_Empresarial B	
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS R ON B.UF_CARTEIRA = R.FILIAL	
                    LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL	

                    WHERE --c.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                    b.grupo_plano not in ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                    AND B.INDBD <> 'CANCELAMENTO' 
                    AND LEFT(DATA,6) = @ANOMES
                    
                    GROUP BY INDBD,	
                        GRUPO_PLANO,
                        left(data,6),
                        C.CANAL_RB,
                        R.REGIONAL_AGRUPADA,
                        LTRIM(RTRIM(B.UF_CARTEIRA)),
                        RIGHT(DATA,2)
                    ) t
                    /*
                    TRUNCATE TABLE TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE_NOVA

                    DECLARE @DATA_2 AS VARCHAR (6)

                    SET @DATA_2 = (SELECT MAX(DATA) FROM TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE)

                    insert into TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE_NOVA
                SELECT 
                    --TIPO_INDICADOR,
                    DATA,
                    INDBD,
                    case 
                                    when GRUPO_PLANO in ('CONTROLE BOLETO','CONTROLE_BOLETO') then 'CONTROLE_BOLETO' 
                                    when GRUPO_PLANO in ('CONTROLE CARTAO','CONTROLE_CARTAO') then 'CONTROLE_CARTAO' 
                                    when GRUPO_PLANO in ('MOVEL','MÓVEL') then 'MOVEL' 
                                    when GRUPO_PLANO in ('OI GALERA PRÉ','PRÉ-PAGO') then 'PRÉ-PAGO' 
                            else GRUPO_PLANO end GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    SEGMENTO,
                    GESTAO GESTÃO,
                    UF,
                    DIA,
                    0 AS COMPROMISSO,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'FORECAST' THEN QTD ELSE 0 END) AS FORECAST,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'META' THEN QTD ELSE 0 END) AS META,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'ORÇAMENTO' THEN QTD ELSE 0 END) AS ORCAMENTO,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'REAL' THEN QTD ELSE 0 END) AS REAL,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'TENDÊNCIA' THEN QTD ELSE 0 END) AS TENDENCIA

                FROM TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE
                /* REMOVENDO AS VISÕES DE FIBRA EM OBRA - SOLICITAÇÃO DO MARIO - 07/05/2019 */
                WHERE GRUPO_PLANO NOT LIKE '%C' AND GRUPO_PLANO NOT LIKE '%EO' AND GRUPO_PLANO NOT LIKE '%SI' AND GRUPO_PLANO NOT LIKE '%EX OBRA'

                GROUP BY 	DATA,
                    INDBD,
                    GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    SEGMENTO,
                    GESTAO,
                    UF,
                    DIA

                UNION ALL

                select 
                    --TIPO_INDICADOR,
                    ANOMES DATA,
                    INDBD,
                    case 
                                    when GRUPO_PLANO in ('CONTROLE BOLETO','CONTROLE_BOLETO') then 'CONTROLE_BOLETO' 
                                    when GRUPO_PLANO in ('CONTROLE CARTAO','CONTROLE_CARTAO') then 'CONTROLE_CARTAO' 
                                    when GRUPO_PLANO in ('MOVEL','MÓVEL') then 'MOVEL' 
                                    when GRUPO_PLANO in ('OI GALERA PRÉ','PRÉ-PAGO') then 'PRÉ-PAGO' 
                            else GRUPO_PLANO end GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    CASE WHEN UNIDADE_NEGOCIO = 'EMPRESARIAL' THEN 'EMP PDV' ELSE UNIDADE_NEGOCIO END AS SEGMENTO,
                    GESTÃO,
                    UF,
                    DIA,
                    SUM(QTD) AS COMPROMISSO,
                    0 FORECAST,
                    0 META,
                    0 ORCAMENTO,
                    0 REAL,
                    0 TENDENCIA

                from TBL_PC_BASEMETA_RELATORIO_MANOEL

                WHERE TIPO_INDICADOR = 'COMPROMISSO'
                AND ANOMES >= @DATA_2

                GROUP BY 	--TIPO_INDICADOR,
                    ANOMES,
                    INDBD,
                    GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    UNIDADE_NEGOCIO,
                    GESTÃO,
                    UF,
                    DIA*/


                    INSERT INTO TBL_PC_TEMPO_PROCEDURES
                    SELECT
                    'SP_PC_CG_IND_Acompanhamento_Diario_Final' AS [PROCEDURE],
                    'FIM' AS INI_FIM,
                    GETDATE() AS DATA_HORA
                    
                    
                    
                    
                    print '--- Commit ---'	
                    PRINT '-----------------------------------------------------------------------------------------'
                    PRINT '-            FIM CARGA SP_PC_CG_IND_Acompanhamento_Diario_Final   MÊS: '+@ANOMES
                    PRINT '-----------------------------------------------------------------------------------------'
                    
                    
                    --COMMIT
                end
				'''

	dados_conexao = (
		"Driver={SQL Server};"
		f"Server={segredos.db_server};"
		f"Database={segredos.db_name};"
		f"UID={segredos.db_user};"
		f"PWD={segredos.db_pass}"
	)
	conexao = pyodbc.connect(dados_conexao)
	print("Conectado ao banco para alterar a procedure - retirar comentários")
	cursor = conexao.cursor()
	cursor.execute(comando_sql)
	conexao.commit()
	conexao.close()
	print('Conexão Fechada')

def ATIVAR_TEND_TABLEAU_teste_Jan22_somenteFibra():
	comando_sql='''
				ALTER PROCEDURE [dbo].[SP_PC_CG_IND_Acompanhamento_Diario_Final] @ANOMES VARCHAR(6)
                WITH RECOMPILE
                AS
                --SET IMPLICIT_TRANSACTIONS ON
                BEGIN-- try 

                    INSERT INTO TBL_PC_TEMPO_PROCEDURES
                    SELECT
                    'SP_PC_CG_IND_Acompanhamento_Diario_Final' AS [PROCEDURE],
                    'INICIO' AS INI_FIM,
                    GETDATE() AS DATA_HORA





                    PRINT '-----------------------------------------------------------------------------------------'
                    PRINT '-            INICIO CARGA SP_PC_CG_IND_Acompanhamento_Diario_Final   MÊS: '+@ANOMES
                    PRINT '-----------------------------------------------------------------------------------------'

                    DECLARE @ANOMES_M1 CHAR(6)
                    SET @ANOMES_M1 = dbo.format_date(DATEADD(MONTH, DATEDIFF(MONTH, 0 , GETDATE()-1)-1,0),'YYYYMM')


                    delete from TBL_IND_VAR_BASEMETA_PORDU where LEFT(data,6) = @ANOMES
                    /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                    --AND TIPO_INDICADOR <> 'TENDÊNCIA'
                    AND NOT (TIPO_INDICADOR = 'TENDÊNCIA' AND GRUPO_PLANO NOT LIKE '%FIBRA')

                    insert into TBL_IND_VAR_BASEMETA_PORDU
                    SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                CASE WHEN GRUPO_PLANO = 'OI GALERA PRÉ' THEN 'PRÉ-PAGO' ELSE GRUPO_PLANO END GRUPO_PLANO,
                                CASE WHEN ((COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                                ELSE C.CANAL_RB END CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.FILIAL,
                                'VAREJO' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BASERESULTADOS AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.FILIAL = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CANAISRELATORIOBERNARDO AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND GRUPO_PLANO NOT IN ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                        AND INDBD <> 'CANCELAMENTO'
                        AND NOT (TIPO_INDICADOR = 'TENDÊNCIA' AND GRUPO_PLANO NOT LIKE '%FIBRA')
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                CASE WHEN GRUPO_PLANO = 'OI GALERA PRÉ' THEN 'PRÉ-PAGO' ELSE GRUPO_PLANO END,
                                DATA,
                                CASE WHEN ((COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                                ELSE C.CANAL_RB END,
                                R.REGIONAL,
                                B.FILIAL
                        UNION ALL
                        SELECT	
                                CASE WHEN (INDBD IN ('VL', 'MIGRACAO')) THEN 'VL' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.FILIAL,
                                'VAREJO' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BASERESULTADOS AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.FILIAL = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CANAISRELATORIOBERNARDO AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND  DATA = @ANOMES
                        AND GRUPO_PLANO IN ('CONECTADO', 'RESIDENCIAL', 'COMPLETO')
                        AND INDBD IN ('VL', 'MIGRACAO')
                        GROUP BY 
                                CASE WHEN (INDBD IN ('VL', 'MIGRACAO')) THEN 'VL' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.FILIAL
                        UNION ALL
                        SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.FILIAL,
                                'VAREJO' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BASERESULTADOS AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.FILIAL = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CANAISRELATORIOBERNARDO AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND GRUPO_PLANO IN ('CONECTADO', 'RESIDENCIAL', 'COMPLETO')
                        AND INDBD IN ('GROSS', 'MIGRACAO')
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.FILIAL
                    UNION ALL
                    SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.UF_CLIENTE FILIAL,
                                'EMP CLI' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BaseResultados_Empresarial AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.UF_CLIENTE = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND INDBD <> 'CANCELAMENTO' 
                        AND NOT (TIPO_INDICADOR = 'TENDÊNCIA' AND GRUPO_PLANO NOT LIKE '%FIBRA')
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.UF_CLIENTE
                        UNION ALL
                        SELECT	
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END INDBD,
                                DATA,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                C.CANAL_RB,
                                R.REGIONAL REGIONAL,
                                B.UF_CARTEIRA FILIAL,
                                'EMP PDV' SEGMENTO,
                                SUM(VALOR) VALOR
                        FROM TBL_RE_BaseResultados_Empresarial AS B
                        LEFT JOIN TBL_RE_DP_REGIONALRELATORIOBERNARDO AS R ON B.UF_CARTEIRA = R.FILIAL
                        LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL
                        /* ************** ATIVAR LINHA ABAIXO PARA TRAVAR TENDÊNCIA ************ */
                        --WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST') AND DATA = @ANOMES
                        /* ************** COMENTAR LINHA ACIMA E ATIVAR LINHA ABAIXO PARA DESTRAVAR TENDÊNCIA ************ */ 
                        WHERE TIPO_INDICADOR IN ('META', 'FORECAST', 'ORÇAMENTO', 'RE-FCST','TENDÊNCIA') AND DATA = @ANOMES
                        AND INDBD <> 'CANCELAMENTO' 
                        AND REGIONAL_AGRUPADA IS NOT NULL
                        AND NOT (TIPO_INDICADOR = 'TENDÊNCIA' AND GRUPO_PLANO NOT LIKE '%FIBRA')
                        GROUP BY 
                                CASE WHEN (INDBD IN ('GROSS', 'MIGRACAO')) THEN 'GROSS' ELSE INDBD END,
                                TIPO_INDICADOR,
                                GRUPO_PLANO,
                                DATA,
                                C.CANAL_RB,
                                R.REGIONAL,
                                B.UF_CARTEIRA



                    DELETE FROM TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE WHERE DATA = @ANOMES
                    AND NOT (TIPO_INDICADOR = 'META' AND DATA IN ('202007', '202008', '202009') AND GRUPO_PLANO IN ('CONTROLE_BOLETO',
                    'CONTROLE_CARTAO', 'PÓS ALONE', 'PÓS OIT', 'PRÉ-PAGO', 'PRÉ-D3', 'MOVEL'))

                    insert into TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE
                    select 
                            TIPO_INDICADOR,
                            data,
                            INDBD,
                            GRUPO_PLANO,
                            CANAL, 
                            REGIONAL,
                            SEGMENTO,
                            CASE 
                                WHEN CANAL IN ('Smart Message','TLV Outros','Outros', 'Outros BRI','Condominios','','NULL','S2S','Outros EMP','S2S EMP','TLV Outros EMP' ) THEN 'Outros Nacionais' 
                                WHEN CANAL IN ('TLV Receptivo','TLV Ativo','TLV Ativo EMP','TLV Receptivo EMP' ) THEN 'TLV' 
                                WHEN CANAL IN ('WEB','WEB EMP') THEN 'WEB' 
                                WHEN CANAL = 'Anteneiros' THEN 'Anteneiros' 
                                WHEN CANAL IN ('TLV Receptivo BRI','TLV Outros BRI','TLV Ativo BRI','TLV BRI' ) THEN 'TLV BRI' 
                                WHEN CANAL LIKE '%EMP%' AND [SEGMENTO] <> 'VAREJO' THEN 'Gestão Regional EMP'
                                WHEN CANAL NOT LIKE '%EMP%' AND [SEGMENTO] <> 'VAREJO' THEN 'Gestão Regional VAR'
                            ELSE 'Gestão Regional' END GESTAO,
                            UF,
                            cast( DIA as varchar(2)) DIA,
                            META_DU QTD


                    from(	SELECT		distinct 
                            TIPO_INDICADOR,
                            a.data,
                            A.indbd INDBD,
                            CASE WHEN A.grupo_plano IN ('OI GALERA PRÉ', 'PRÉ-PAGO') THEN 'PRÉ-PAGO' ELSE A.grupo_plano END GRUPO_PLANO,
                            A.canal_rb CANAL, 
                            REG.REGIONAL_AGRUPADA REGIONAL,
                            SEGMENTO,
                            LTRIM(RTRIM(A.FILIAL)) UF,
                            '' AS DIA,
                            a.VALOR META_DU
                            -- ,A.VALOR
                            --,DU_MES_PRODUTO.VALOR DU_MES_PRODUTO
                            -- ,DU_MES.VALOR DU_MES
                            -- ,P.VALOR P
                        
                    FROM TBL_IND_VAR_BASEMETA_PORDU A
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS REG ON A.FILIAL = REG.FILIAL
                                    
                    WHERE --a.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                                a.grupo_plano not in (	'CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS',
                                                            'TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS',
                                                            'BANDA LARGA FIBRA C','BANDA LARGA FIBRA EO','BANDA LARGA FIBRA EX OBRA','BANDA LARGA FIBRA SI','FIBRA C',
                                                            'FIBRA EO','FIBRA EX OBRA','FIBRA SI','FIXO FIBRA C','FIXO FIBRA EO','FIXO FIBRA EX OBRA','FIXO FIBRA SI',
                                                            'OI TV FIBRA C','OI TV FIBRA EO','OI TV FIBRA EX OBRA','OI TV FIBRA SI','OIT COMERCIAL',
                                                            'OIT CONECTADO','OIT SOLUCAO COMPLETA','VADA')
                            AND A.TIPO_INDICADOR = 'TENDÊNCIA'
                            AND A.DATA = @ANOMES

                            UNION ALL


                    SELECT		distinct 
                            TIPO_INDICADOR,
                            a.data,
                            A.indbd INDBD,
                            CASE WHEN A.grupo_plano IN ('OI GALERA PRÉ', 'PRÉ-PAGO') THEN 'PRÉ-PAGO' ELSE A.grupo_plano END GRUPO_PLANO,
                            A.canal_rb CANAL, 
                            REG.REGIONAL_AGRUPADA REGIONAL,
                            SEGMENTO,
                            LTRIM(RTRIM(A.FILIAL)) UF,
                            RIGHT(C.ANOMESDIA,2) AS DIA,
                            case 
                                when A.VALOR = 0 then 0
                                when DU_MES.VALOR > 0 or DU_MES.VALOR is not null then (A.VALOR/DU_MES.VALOR)*C.VALOR
                                when DU_MES_PRODUTO.VALOR > 0 or DU_MES_PRODUTO.VALOR is not null then (A.VALOR/DU_MES_PRODUTO.VALOR)*p.VALOR 
                            else	0 end META_DU
                            -- ,A.VALOR
                            --,DU_MES_PRODUTO.VALOR DU_MES_PRODUTO
                            -- ,DU_MES.VALOR DU_MES
                            -- ,P.VALOR P
                        
                    FROM TBL_IND_VAR_BASEMETA_PORDU A
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS REG ON A.FILIAL = REG.FILIAL

                    /* TOTAL DU POR REGIONAL - CANAL */
                    left JOIN (
                                    SELECT 
                                        * 
                                    FROM TBL_IND_VAR_DU_SUMARIZADO
                                    where ANOMES = @ANOMES
                            ) DU_MES ON A.INDBD = DU_MES.INDBD AND 
                                                        --a.GRUPO_PLANO = DU_MES.PRODUTO_NOVO AND
                                                        case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = DU_MES.PRODUTO_NOVO AND
                                                        CASE WHEN A.CANAL_RB = 'Smart Message' THEN 'TLV Outros' 
                                                        ELSE A.CANAL_RB END = DU_MES.CANAL AND
                                                        A.REGIONAL = DU_MES.REGIONAL AND
                                                        A.DATA = DU_MES.ANOMES

                    /* DU POR DIA PARA REGIONAL - CANAL */
                    left JOIN (SELECT ANOMESDIA,
                                    left(ANOMESDIA,6) ANOMES,
                                    INDBD,
                                    CASE WHEN DU_DIA.PRODUTO IN ('OI GALERA PRÉ', 'PRÉ-PAGO') THEN 'PRÉ-PAGO' ELSE DU_DIA.PRODUTO END PRODUTO_NOVO,
                                    CANAL,
                                    REGIONAL,
                                    valor
                                FROM TBL_pc_du AS DU_DIA
                                            where left(ANOMESDIA,6) = @ANOMES
                                ) C ON 
                                        A.INDBD = C.INDBD AND 
                                        case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = C.PRODUTO_NOVO AND
                                        CASE WHEN A.CANAL_RB = 'Smart Message' THEN 'TLV Outros' ELSE
                                        A.CANAL_RB END = C.CANAL AND
                                        A.REGIONAL = C.REGIONAL and
                                        A.DATA = C.ANOMES
                    
                    /* TOTAL DU POR PRODUTOS */
                    left JOIN (
                                    SELECT 
                                            LEFT(ANOMESDIA,6) ANOMES,
                                            INDBD,
                                            PRODUTO_NOVO,
                                            SUM(cast(rtrim(ltrim(replace(VALOR,',','.'))) as float)) VALOR 
                                    FROM TBL_PC_DU_PRODUTO
                                    where left(ANOMESDIA,6) = @ANOMES
                                    group by 
                                            LEFT(ANOMESDIA,6),
                                            INDBD,
                                            PRODUTO_NOVO
                            ) DU_MES_PRODUTO ON	A.INDBD = DU_MES_PRODUTO.INDBD AND 
                                                    case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = DU_MES_PRODUTO.PRODUTO_NOVO AND
                                                    A.DATA = DU_MES_PRODUTO.ANOMES

                    /* DU POR DIA PARA PRODUTO */
                    left JOIN (SELECT 
                                        ANOMESDIA,
                                        LEFT(ANOMESDIA,6) ANOMES,
                                        INDBD,
                                        PRODUTO_NOVO,
                                        SUM(cast(rtrim(ltrim(replace(VALOR,',','.'))) as float)) VALOR 
                                FROM TBL_PC_DU_PRODUTO
                                where left(ANOMESDIA,6) = @ANOMES
                                group by 
                                        ANOMESDIA,
                                        LEFT(ANOMESDIA,6),
                                        INDBD,
                                        PRODUTO_NOVO
                                ) P ON 
                                        A.INDBD = P.INDBD AND 
                                        case when A.grupo_plano = 'NOVA FIBRA' then 'FIBRA' ELSE A.GRUPO_PLANO END = P.PRODUTO_NOVO AND
                                        A.DATA = P.ANOMES
                    
                                    
                    WHERE --a.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                                a.grupo_plano not in (	'CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS',
                                                            'TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS',
                                                            'BANDA LARGA FIBRA C','BANDA LARGA FIBRA EO','BANDA LARGA FIBRA EX OBRA','BANDA LARGA FIBRA SI','FIBRA C',
                                                            'FIBRA EO','FIBRA EX OBRA','FIBRA SI','FIXO FIBRA C','FIXO FIBRA EO','FIXO FIBRA EX OBRA','FIXO FIBRA SI',
                                                            'OI TV FIBRA C','OI TV FIBRA EO','OI TV FIBRA EX OBRA','OI TV FIBRA SI','OIT COMERCIAL',
                                                            'OIT CONECTADO','OIT SOLUCAO COMPLETA','VADA')
                            AND A.TIPO_INDICADOR <> 'TENDÊNCIA'
                            AND A.DATA = @ANOMES
                            AND NOT (TIPO_INDICADOR = 'META' AND DATA IN ('202007', '202008', '202009') AND GRUPO_PLANO IN ('CONTROLE_BOLETO',
                            'CONTROLE_CARTAO', 'PÓS ALONE', 'PÓS OIT', 'PRÉ-PAGO', 'PRÉ-D3', 'MOVEL'))

                    union all

                    SELECT  'REAL' TIPO_INDICADOR,
                            left(data,6) anomes,
                            INDBD,	
                        GRUPO_PLANO,
                        CASE WHEN ((B.COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                        ELSE C.CANAL_RB END canal,
                        R.REGIONAL_AGRUPADA regional,
                        'VAREJO' SEGMENTO,
                        LTRIM(RTRIM(B.FILIAL)) AS UF,
                        RIGHT(DATA,2) AS DIA,
                        SUM(B.QTD) AS valor

                    FROM TBL_RE_BaseResultadoDiario B	
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS R ON B.FILIAL = R.FILIAL	
                    LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo AS C ON B.CANAL_FINAL = C.CANAL_FINAL	

                    WHERE --c.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                    b.grupo_plano not in ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                    AND B.INDBD <> 'CANCELAMENTO' 
                    AND LEFT(DATA,6) = @ANOMES
                    
                    GROUP BY INDBD,	
                        GRUPO_PLANO,
                        left(data,6),
                        CASE WHEN ((B.COD_SAP = '1042793') AND (GRUPO_PLANO = 'CONTROLE_BOLETO')) THEN 'Smart Message'
                        ELSE C.CANAL_RB END,
                        R.REGIONAL_AGRUPADA,
                    
                        LTRIM(RTRIM(B.FILIAL)),
                        RIGHT(DATA,2)

                    union all

                    SELECT  'REAL' TIPO_INDICADOR,
                            left(data,6) anomes,
                            INDBD,	
                        GRUPO_PLANO,
                        C.CANAL_RB canal,
                        R.REGIONAL_AGRUPADA regional,
                        'EMP CLI' SEGMENTO,
                        LTRIM(RTRIM(B.UF_CLIENTE)) AS UF,
                        RIGHT(DATA,2) AS DIA,
                        SUM(B.QTD) AS valor

                    FROM TBL_RE_BaseResultadoDiario_Empresarial B	
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS R ON B.UF_CLIENTE = R.FILIAL	
                    LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL	

                    WHERE --c.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                    b.grupo_plano not in ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                    AND B.INDBD <> 'CANCELAMENTO' 
                    AND LEFT(DATA,6)  = @ANOMES
                    
                    GROUP BY INDBD,	
                        GRUPO_PLANO,
                        left(data,6),
                        C.CANAL_RB,
                        R.REGIONAL_AGRUPADA,
                    
                        LTRIM(RTRIM(B.UF_CLIENTE)),
                        RIGHT(DATA,2)

                    union all

                    SELECT  'REAL' TIPO_INDICADOR,
                            left(data,6) anomes,
                            INDBD,	
                        GRUPO_PLANO,
                        C.CANAL_RB canal,
                        R.REGIONAL_AGRUPADA regional,
                        'EMP PDV' SEGMENTO,
                        LTRIM(RTRIM(B.UF_CARTEIRA)) AS UF,
                        RIGHT(DATA,2) AS DIA,
                        SUM(B.QTD) AS valor
                    FROM TBL_RE_BaseResultadoDiario_Empresarial B	
                    LEFT JOIN TBL_RE_DP_RegionalRelatorioBernardo AS R ON B.UF_CARTEIRA = R.FILIAL	
                    LEFT JOIN TBL_RE_DP_CanaisRelatorioBernardo_Empresarial AS C ON B.CANAL_FINAL = C.CANAL_FINAL	

                    WHERE --c.canal_rb not in ('outros bri', 'tlv ativo bri', 'tlv receptivo bri', 'tlv outros bri','Anteneiros','tlv bri','bri') and
                    b.grupo_plano not in ('CONECTADO', 'RESIDENCIAL', 'COMPLETO','1P','2P BL','3G','CONTROLE','CONTROLE_BOLETO-M1','OI TV PRE-PAGO','PACOTE_DADOS','TV ANTENEIROS','PÓS OCT','PÓS OCT PLANOS','CONTROLE_VOZ','BL+POS','PRÉ-M1','PÓS TOTAL','PÓS TOTAL PLANOS')
                    AND B.INDBD <> 'CANCELAMENTO' 
                    AND LEFT(DATA,6) = @ANOMES
                    
                    GROUP BY INDBD,	
                        GRUPO_PLANO,
                        left(data,6),
                        C.CANAL_RB,
                        R.REGIONAL_AGRUPADA,
                        LTRIM(RTRIM(B.UF_CARTEIRA)),
                        RIGHT(DATA,2)
                    ) t
                    /*
                    TRUNCATE TABLE TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE_NOVA

                    DECLARE @DATA_2 AS VARCHAR (6)

                    SET @DATA_2 = (SELECT MAX(DATA) FROM TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE)

                    insert into TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE_NOVA
                SELECT 
                    --TIPO_INDICADOR,
                    DATA,
                    INDBD,
                    case 
                                    when GRUPO_PLANO in ('CONTROLE BOLETO','CONTROLE_BOLETO') then 'CONTROLE_BOLETO' 
                                    when GRUPO_PLANO in ('CONTROLE CARTAO','CONTROLE_CARTAO') then 'CONTROLE_CARTAO' 
                                    when GRUPO_PLANO in ('MOVEL','MÓVEL') then 'MOVEL' 
                                    when GRUPO_PLANO in ('OI GALERA PRÉ','PRÉ-PAGO') then 'PRÉ-PAGO' 
                            else GRUPO_PLANO end GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    SEGMENTO,
                    GESTAO GESTÃO,
                    UF,
                    DIA,
                    0 AS COMPROMISSO,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'FORECAST' THEN QTD ELSE 0 END) AS FORECAST,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'META' THEN QTD ELSE 0 END) AS META,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'ORÇAMENTO' THEN QTD ELSE 0 END) AS ORCAMENTO,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'REAL' THEN QTD ELSE 0 END) AS REAL,
                    SUM(CASE WHEN  TIPO_INDICADOR = 'TENDÊNCIA' THEN QTD ELSE 0 END) AS TENDENCIA

                FROM TBL_IND_VAR_ACOMPANHAMENTO_RESIDENCIAL_BASE
                /* REMOVENDO AS VISÕES DE FIBRA EM OBRA - SOLICITAÇÃO DO MARIO - 07/05/2019 */
                WHERE GRUPO_PLANO NOT LIKE '%C' AND GRUPO_PLANO NOT LIKE '%EO' AND GRUPO_PLANO NOT LIKE '%SI' AND GRUPO_PLANO NOT LIKE '%EX OBRA'

                GROUP BY 	DATA,
                    INDBD,
                    GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    SEGMENTO,
                    GESTAO,
                    UF,
                    DIA

                UNION ALL

                select 
                    --TIPO_INDICADOR,
                    ANOMES DATA,
                    INDBD,
                    case 
                                    when GRUPO_PLANO in ('CONTROLE BOLETO','CONTROLE_BOLETO') then 'CONTROLE_BOLETO' 
                                    when GRUPO_PLANO in ('CONTROLE CARTAO','CONTROLE_CARTAO') then 'CONTROLE_CARTAO' 
                                    when GRUPO_PLANO in ('MOVEL','MÓVEL') then 'MOVEL' 
                                    when GRUPO_PLANO in ('OI GALERA PRÉ','PRÉ-PAGO') then 'PRÉ-PAGO' 
                            else GRUPO_PLANO end GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    CASE WHEN UNIDADE_NEGOCIO = 'EMPRESARIAL' THEN 'EMP PDV' ELSE UNIDADE_NEGOCIO END AS SEGMENTO,
                    GESTÃO,
                    UF,
                    DIA,
                    SUM(QTD) AS COMPROMISSO,
                    0 FORECAST,
                    0 META,
                    0 ORCAMENTO,
                    0 REAL,
                    0 TENDENCIA

                from TBL_PC_BASEMETA_RELATORIO_MANOEL

                WHERE TIPO_INDICADOR = 'COMPROMISSO'
                AND ANOMES >= @DATA_2

                GROUP BY 	--TIPO_INDICADOR,
                    ANOMES,
                    INDBD,
                    GRUPO_PLANO,
                    CANAL,
                    REGIONAL,
                    UNIDADE_NEGOCIO,
                    GESTÃO,
                    UF,
                    DIA*/

                    INSERT INTO TBL_PC_TEMPO_PROCEDURES
                    SELECT
                    'SP_PC_CG_IND_Acompanhamento_Diario_Final' AS [PROCEDURE],
                    'FIM' AS INI_FIM,
                    GETDATE() AS DATA_HORA


                    print '--- Commit ---'	
                    PRINT '-----------------------------------------------------------------------------------------'
                    PRINT '-            FIM CARGA SP_PC_CG_IND_Acompanhamento_Diario_Final   MÊS: '+@ANOMES
                    PRINT '-----------------------------------------------------------------------------------------'
                    
                    
                    --COMMIT
                end
				'''

	dados_conexao = (
		"Driver={SQL Server};"
		f"Server={segredos.db_server};"
		f"Database={segredos.db_name};"
		f"UID={segredos.db_user};"
		f"PWD={segredos.db_pass}"
	)
	conexao = pyodbc.connect(dados_conexao)
	print("Conectado ao banco para alterar a procedure - colocar comentários")
	cursor = conexao.cursor()
	cursor.execute(comando_sql)
	conexao.commit()
	conexao.close()
	print('Conexão Fechada')

def executa_procedure_sql(nome_procedure, param):
   
    dados_conexao = (
        "Driver={SQL Server};"
        f"Server={segredos.db_server};"
        f"Database={segredos.db_name};"
        f"UID={segredos.db_user};"
        f"PWD={segredos.db_pass}"
    )
    conexao = pyodbc.connect(dados_conexao)
    print('\x1b[1;33;42m' + 'Conexão realizada ao banco de dados' + '\x1b[0m')

    cursor = conexao.cursor()
    
    #executa procedure
    inicio_procedure = datetime.today()
    print('\x1b[1;33;44m' + f'Executando a Procedure {nome_procedure} para o parâmetro: {param} '+ '\x1b[0m')
    print(f'Iniciando execução em: {inicio_procedure}')
    cursor.execute(f'SET NOCOUNT ON; EXEC {nome_procedure}  {param}')
    conexao.commit()
    fim_procedure = datetime.today()
    print(f"Procedure executada em {fim_procedure - inicio_procedure} tempo")
    
    conexao.close()
    print('\x1b[1;33;41m' + 'Conexão Fechada'+ '\x1b[0m')



param = datetime.today().strftime('%Y%m')

ATIVAR_TEND_TABLEAU_teste_Jan22()
executa_procedure_sql('SP_PC_BASES_SHAREPOINT',param)
ATIVAR_TEND_TABLEAU_teste_Jan22_somenteFibra()