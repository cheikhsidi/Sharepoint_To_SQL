/****** Object:  StoredProcedure [etl].[SharePoint_PhysicalLocations]    Script Date: 8/10/2020 12:34:33 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



/*select * from frps.PhysicalLocations
select * from ETL.OperatingUnit*/

CREATE proc [etl].[SharePoint_PhysicalLocations] as 

insert into etl.SharePoint_BatchCheck
select getdate()
	, 'Starting'
	, (select min(sp_id) from etl.OperatingUnit)
	

begin try

    SELECT    
       SUBSTRING(
        (
            SELECT ',' + a.DealName  AS [text()]
            FROM frps.PhysicalLocations_Deal as a 
            JOIN etl.PhysicalLocations as  b ON a.PhysicalLocation_ID = b.SP_ID
            ORDER BY a.PhysicalLocation_ID
            FOR XML PATH ('')
        ), 2, 1000) as  Deal



declare @Result as nvarchar(400)
set @Result = 'Nothing Updated or Loaded!!!'

/*Critical errror - more than 2 records in etl table*/
if
	(select count(*) from etl.PhysicalLocations) > 1
begin
	insert into etl.PhysicalLocations_error
	select * from etl.PhysicalLocations

	set @Result = 'Critital Error more than one record in ETL'
end  

/*SharePoint List is updated*/
else if
	(select count(*) from frps.PhysicalLocations where ID in (select SP_ID from etl.PhysicalLocations)) = 1

begin
	
	insert into frps.PhysicalLocations_history
	select 
     SUBSTRING(
        (
            SELECT ',' + a.DealName  AS [text()]
            FROM frps.PhysicalLocations_Deal as a 
            JOIN etl.PhysicalLocations as  b ON a.PhysicalLocation_ID = b.SP_ID
            ORDER BY a.PhysicalLocation_ID
            FOR XML PATH ('')
        ), 2, 1000),
    SUBSTRING(
        (
            SELECT ',' + a.BusinessUnit  AS [text()]
            FROM frps.PhysicalLocations_BusinessUnit as a 
            JOIN etl.PhysicalLocations as  b ON a.PhysicalLocation_ID = b.SP_ID
            ORDER BY a.PhysicalLocation_ID
            FOR XML PATH ('')
        ), 2, 1000),
    a.*, 
    'Update'
	from frps.PhysicalLocations as a 
	join etl.PhysicalLocations as b
	on a.ID = b.SP_ID
	--where ID in (select SP_ID from etl.PhysicalLocations) 

	update frps.PhysicalLocations
	set [Address Line] = (select AddressLine from etl.PhysicalLocations)
		, [Address Line 2] = (select [AddressLine2] from etl.PhysicalLocations)
        , City = (Select City from etl.PhysicalLocations)
        , [State/Territory] = (Select [State/Territory] from etl.PhysicalLocations )
        , Zip = (Select Zip from etl.PhysicalLocations )
        , [Mailing Address] = (Select MailingAddress from etl.PhysicalLocations )
        , [Site Description] = (Select SiteDescription from etl.PhysicalLocations )
        , ActiveInactive = (Select ActiveInactive from etl.PhysicalLocations )
        , BU_Code = (Select BU_Code from etl.PhysicalLocations )
		, [ModifiedBy] = (select ModifiedBy from etl.PhysicalLocations)	
		, ModifiedDate = (select ModifiedDate from etl.PhysicalLocations)
        , EffectiveDate = (select EffectiveDate from etl.PhysicalLocations)
        , ExpirationDate = (select ExpirationDate from etl.PhysicalLocations)
	where ID in (select SP_ID from etl.PhysicalLocations) 

   
    delete from frps.PhysicalLocations_Deal where ID in (select SP_ID from etl.PhysicalLocations)
    insert into frps.PhysicalLocations_Deal
    select OtherId, cs.Value
    from etl.PhysicalLocations
    cross apply (select Deal from dbo.Split(t.Data,',') ) cs

    -- update frps.PhysicalLocations_Deal
	-- set [DealName] = (select AddressLine from etl.PhysicalLocations)
	-- where ID in (select SP_ID from etl.PhysicalLocations) 
	set @Result = 'Updated Record'
end

/*New Record added that replaces another*/

--else if 
--	(select count(*) from frps.PhysicalLocations where ID in (select SP_ID from etl.PhysicalLocations)) = 0
--	and 
--	(select count(*)
--		from frps.PhysicalLocations as a
--		join etl.PhysicalLocations as b
--			on a.Name = b.Name) > 0
--	and
--	(select count(*) from frps.PhysicalLocations_history where ID in (select SP_ID from etl.PhysicalLocations)) = 0
--begin
--	/*Move previous record to history*/
--	insert into frps.PhysicalLocations_history
--	select a.*, 'Replace', b.SP_ID 
--	from frps.PhysicalLocations as a
--	join etl.PhysicalLocations as b
--	on a.Name = b.Name
--	/*where ID in (select a.ID
--					from frps.PhysicalLocations as a
--					join etl.PhysicalLocations as b
--						on a.Name = b.Name) */

--	--------------------------------------------------------------------
--	---check from Hisotrical records with Active SPs that need to be updated
--	Update frps.PhysicalLocations_history
--	set Active_SP = b.SP_ID
--	from frps.PhysicalLocations as a
--	join etl.PhysicalLocations as b
--		on a.Name = b.Name
--	--where a.Code in (Select Code from etl.DataSource)


--	/*Drop previous record*/
--	delete frps.PhysicalLocations
--	where ID in (select a.ID
--					from frps.PhysicalLocations as a
--					join etl.PhysicalLocations as b
--						on a.Name = b.Name) 
--	/*Insert change*/
--	insert into frps.PhysicalLocations
--	select BusinessUnitID
--		, Name
--		, ModifiedBy
--		, ModifiedDate
--	from etl.PhysicalLocations

--	set @Result = 'Record has been replaced'
--end

/*New Record added */
else if
	(select count(*) from frps.PhysicalLocations where ID in (select SP_ID from etl.PhysicalLocations)) = 0
		and
	(select count(*) from frps.PhysicalLocations_history where ID in (select SP_ID from etl.PhysicalLocations)) = 0
begin
	
	insert into frps.PhysicalLocations
	select Sp_ID 
		, AddressLine
		, [AddressLine2]
        , City 
        , [State/Territory]
        , Zip
        , MailingAddress
        , SiteDescription
        , ActiveInactive
        , BU_Code
		, ModifiedBy
		, ModifiedDate
        , EffectiveDate
        , ExpirationDate
	from etl.PhysicalLocations

	set @Result = 'Loaded Record'

end

/*Update to a record in history*/
--else if 
--	(select count(*) from frps.PhysicalLocations where ID in (select SP_ID from etl.PhysicalLocations)) = 0
--	and 
--	(select count(*) from frps.PhysicalLocations_history where ID in (select SP_ID from etl.PhysicalLocations)) > 0
--begin
--	insert into frps.PhysicalLocations_history
--	select a.*, 'Update', b.SP_ID
--	from frps.PhysicalLocations as a
--    join etl.PhysicalLocations as b
--    on a.ID = b.SP_ID
	
--	update frps.PhysicalLocations
--	set BusinessUnitID = (select BusinessUnitID from etl.PhysicalLocations)
--		, Name = (select Name from etl.PhysicalLocations)
--		, ModifiedBy = (select ModifiedBy from etl.PhysicalLocations)	
--		, ModifiedDate = (select ModifiedDate from etl.PhysicalLocations)
--	where ID in (select Active_SP from frps.PhysicalLocations_history) 

--	set @Result = 'Updated The New record.'
	
--end


truncate table etl.PhysicalLocations

insert into etl.SharePoint_BatchCheck
select getdate()
	, @Result
	, (select min(sp_id) from etl.PhysicalLocations)

select @Result as Result
return



end try
begin catch
	


	set @Result = 'FAILURE - ' + ERROR_MESSAGE()

	insert into etl.SharePoint_BatchCheck
	select getdate()
	, @Result
	, (select min(sp_id) from etl.PhysicalLocations)

	select @Result as Result
	return
	
end catch





GO


