# Excel To XML Converter

This python script is used to convert an excel file that contains all of the University of La Verne's colleges and corresponding departments into xml records that will be used to uplaod into the Esploro platform for user/profile manageemnt. Python script will iterate through all rows in excel file to find each college name and pair them with corresponding departments. 

Examlpe XML Heirarchy (based on Esploro documentation):

<unit>
  <unitData>
    <collegeName></collegeName>
  </unitData>
  <subunits>
    <unit>
      <unitData>
        <departmentName></departmentName>
      </unitData>
    </unit>
  </subunits>
 </unit>
