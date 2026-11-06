SELECT
SiteId,
NameLast,
CONCAT(LEFT(NameLast, 1), '.') AS NameLastInitial,
NameFirst,
NameFirstPreferred,
JobCategoryDescr
FROM vwHREmploymentList
WHERE
EmploymentStatusIsActive = 1
  -- AND NameLast = 'Cooper'
  AND (
        (SiteId = 540 AND JobCategoryDescr <> 'Instructional Assistance')
     OR (SiteId = 570 AND JobCategoryDescr = 'Clerical')
     OR (SiteId = 640 AND JobCategoryDescr = 'Information Technology')
  )
ORDER BY NameLast;

