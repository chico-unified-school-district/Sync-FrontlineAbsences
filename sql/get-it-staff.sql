select empid,SiteId,NameLast,NameFirst,NameFirstPreferred,JobCategoryDescr
 from vwHREmploymentList
where
EmploymentStatusIsActive = 1
AND (
 SiteId = 540 AND JobCategoryDescr <> 'Instructional Assistance'
 ) OR (
 SiteId = 640 AND JobCategoryDescr = 'Information Technology'
 )
ORDER BY NameLast;