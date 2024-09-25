-- FUNCTION: public.is_valid_date(text)

-- DROP FUNCTION public.is_valid_date(text);

CREATE OR REPLACE FUNCTION public.is_valid_date(
	text)
    RETURNS boolean
    LANGUAGE 'plpgsql'

    COST 100
    IMMUTABLE 
AS $BODY$
begin
  return case when $1::date is null then false else true end;
exception when others then
  return false;
end;$BODY$;

ALTER FUNCTION public.is_valid_date(text)
    OWNER TO pythonuser;
   
-- Table: public.raw_CovidTrackerGantt

DROP TABLE IF EXISTS public."raw_CovidTrackerGantt";

/*
Index(['CountryName', 'CountryCode', 'Date', 'Name', 'Category_x',
       'CategoryJoin', 'ShortCode', 'ShortDescription', 'Description',
       'Measurement', 'Category_y', 'Type', 'ActivityLong', 'Activity',
       'ActivityGroup', 'prev_name', 'prev_category', 'prev_activity',
       'prev_date', 'newtask', 'startdate', 'enddate'],
      dtype='object')
*/

CREATE TABLE public."raw_CovidTrackerGantt"
(
    index bigint,
    "CountryName" text COLLATE pg_catalog."default",
    "CountryCode" text COLLATE pg_catalog."default",
    "Date" text COLLATE pg_catalog."default",
    "RegionName" text COLLATE pg_catalog."default",
    "RegionCode" text COLLATE pg_catalog."default",
    "Name" text COLLATE pg_catalog."default",
    "Category_x" text COLLATE pg_catalog."default",
    "CategoryJoin" text COLLATE pg_catalog."default",
    "ShortCode" text COLLATE pg_catalog."default",
    "ShortDescription" text COLLATE pg_catalog."default",
    "Description" text COLLATE pg_catalog."default",
    "Measurement" text COLLATE pg_catalog."default",
    "Category_y" text COLLATE pg_catalog."default",
    "Type" text COLLATE pg_catalog."default",
    "ActivityLong" text COLLATE pg_catalog."default",
    "Activity" text COLLATE pg_catalog."default",
    "ActivityGroup" text COLLATE pg_catalog."default",
    next_name text COLLATE pg_catalog."default",
    next_category text COLLATE pg_catalog."default",
    next_activity text COLLATE pg_catalog."default",
    next_countrycode text COLLATE pg_catalog."default",
    next_regioncode text COLLATE pg_catalog."default",
    next_date text COLLATE pg_catalog."default",
    datediff text COLLATE pg_catalog."default",
    newtask text COLLATE pg_catalog."default",
    startdate text COLLATE pg_catalog."default",
    enddate text COLLATE pg_catalog."default"
)
WITH (
    OIDS = FALSE
)
TABLESPACE pg_default;

ALTER TABLE public."raw_CovidTrackerGantt"
    OWNER to pythonuser;
