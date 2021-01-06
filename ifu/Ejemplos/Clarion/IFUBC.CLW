  MEMBER('IFU.clw')

  PRAGMA('define(init_priority=>3)')

  MAP
    MODULE('IFUBC0.CLW')
IFUBC0:DctInit             PROCEDURE
IFUBC0:DctKill             PROCEDURE
IFUBC0:FilesInit           PROCEDURE
    END
  END

DctInit PROCEDURE
  CODE
  IFUBC0:DctInit
  IFUBC0:FilesInit


DctKill PROCEDURE
  CODE
  IFUBC0:DctKill

