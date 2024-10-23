WITH Src AS (
    SELECT
        CASE
            -- info plus:
            WHEN value LIKE '%spContactsTabListExtended%' THEN 'Kontakty - Info plus'
            WHEN value LIKE '%spOffersTabListExtended%' THEN 'Nabídky - Info plus'
            WHEN value LIKE '%spBusinessResultTabGroupByCustomerListExtended%' THEN 'Obchodní výsledky - Info plus'
            WHEN value LIKE '%spBusinessResultTabGroupByCustomerBuListExtended%' THEN 'Obchodní výsledky - Seskupeno podle BU - Info plus'
            WHEN value LIKE '%spBusinessResultTabGroupByCustomerBuGckListExtended%' THEN 'Obchodní výsledky - Seskupeno podle BU + GBK - Info plus'
            WHEN value LIKE '%spBusinessResultTabUngroupedListExtended%' THEN 'Obchodní výsledky - Neseskupeno - Info plus'
            WHEN value LIKE '%fnServiceOrderTabListDataSourceExtended%' THEN 'Servis - Info plus'
            WHEN value LIKE '%spBusinessResultTabGroupByCustomerList%' THEN 'Obchodní výsledky'
            WHEN value LIKE '%spBusinessResultTabGroupByCustomerBuList%' THEN 'Obchodní výsledky - Seskupeno podle BU'
            WHEN value LIKE '%spBusinessResultTabGroupByCustomerBuGckList%' THEN 'Obchodní výsledky - Seskupeno podle BU + GBK'

            -- standart:
            WHEN value LIKE '%spContactsTabList%' THEN 'Kontakty'
            WHEN value LIKE '%spOffersTabList%' THEN 'Nabídky'
            WHEN value LIKE '%fnReturnOrder2%' THEN 'Vratky / Log. Reklam.'
            WHEN value LIKE '%spBusinessResultTabUngroupedList%' THEN 'Obchodní výsledky - Neseskupeno'
            WHEN value LIKE '%spDebtToDateTimePartitionedTabList%' THEN 'Pohledávky - Rozdělené po časových periodách'
            WHEN value LIKE '%spDebtToDateDetailedTabList%' THEN 'Pohledávky - Detailmí zobrazení'
            WHEN value LIKE '%spEBusinessStatsTabList%' THEN 'E-business'
            WHEN value LIKE '%fnServiceOrderTabListDataSource%' THEN 'Servis'
            WHEN value LIKE '%spServiceRateTabList%' THEN 'Servis - Sazby'
            WHEN value LIKE '%fnServiceRateTabListDataSource%' THEN 'Servis - Servisní výjezdy k zakázce'
            WHEN value LIKE '%fnOffersTabDataSource%' THEN 'Načíst změněný záznam pro horní mřížku'


            -- dolni zalozky - info plus:
            WHEN value LIKE '%spCustomerOffersTabListExtended%' THEN 'Zákazníci - Nabídky - Info plus'
            WHEN value LIKE '%spCustomersTabListExtended%' THEN 'Zákazníci - Info plus'
            WHEN value LIKE '%spCustomerBusinessResultTabGroupByCustomerBuGckListExtended%' THEN 'Zákazníci - Obchodní výsledky - Seskupeno podle BU + GBK - Info plus'
            WHEN value LIKE '%spCustomerBusinessResultByMonthTabGroupByBuGckListExtended%' THEN 'Zákazníci - Měsíční obchodní výsledky - Seskupeno podle BU + GBK - Info plus'
            WHEN value LIKE '%spCustomerBusinessResultByMonthTabUngroupedListExtended%' THEN 'Zákazníci - Měsíční obchodní výsledky - Neseskupeno - Info plus'
            WHEN value LIKE '%spCustomerBusinessResultTabUngroupedListExtended%' THEN 'Zákazníci - Ochodní výsledky - Neseskupeno - Info plus'

            -- dolni zalozky - standart:
            WHEN value LIKE '%spCustomersTabList%' THEN 'Zákazníci'
            WHEN value LIKE '%fnCustomerTabDetail%' THEN 'Zákazníci'
            WHEN value LIKE '%spCustomerOffersTabList%' THEN 'Zákazníci - Nabídky'
            WHEN value LIKE '%spCustomerEBusinessTabList%' THEN 'Zákazníci - E-business'
            WHEN value LIKE '%spCustomerDebtToDateList%' THEN 'Zákazníci - Pohledávky'
            WHEN value LIKE '%CustomerSalesTransparency%' THEN 'Zákazníci - Sales Transparency'
            WHEN value LIKE '%CustomerSalesPlanLock%' THEN 'Zákazníci - Sales planning'
            WHEN value LIKE '%spCustomerBusinessResultByMonthTabGroupByBuGckList%' THEN 'Zákazníci - Měsíční obchodní výsledky - Seskupeno podle BU + GBK'
            WHEN value LIKE '%spCustomerBusinessResultByMonthTabUngroupedList%' THEN 'Zákazníci - Měsíční obchodní výsledky - Neseskupeno'
            WHEN value LIKE '%spCustomerBusinessResultTabUngroupedList%' THEN 'Zákazníci - Obchodní výsledky - Neseskupeno'
            WHEN value LIKE '%spCustomerBusinessResultTabGroupByCustomerBuGckList%' THEN 'Zákazníci - Obchodní výsledky - Seskupeno podle BU + GBK'

            ELSE 'Neznámé'
            END AS MappedValue,

        value AS Value1,
        *
    FROM [importdb].[dbo].[_SQLServerProfiler_temp_source]
    WHERE
        (
            value LIKE '%sp%Tab%' OR value LIKE '%fn%Tab%'
                OR value LIKE '%fnReturnOrder2%'
                OR value LIKE '%CustomerSalesPlanLock%'
                OR value LIKE '%spCustomerDebtToDateList%'
                OR value LIKE '%CustomerSalesTransparency%'
            )
      AND LoginName NOT LIKE '%kahleova%'
      AND value NOT LIKE '%dbo.fnKstTables(%'
      AND value NOT LIKE '%dbo.fnKstAdministrationAssignedTables(%'
      AND ApplicationName = 'Core Microsoft SqlClient Data Provider'
)
SELECT DISTINCT
    value AS "NazevProcedury", MappedValue,
    NTUserName,
    COUNT(*) OVER (PARTITION BY value, NTUserName) AS "Count"
FROM  src
ORDER BY
    4 desc

