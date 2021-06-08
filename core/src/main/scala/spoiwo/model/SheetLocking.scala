package spoiwo.model

object SheetLocking {

  lazy val AllUnlocked: SheetLocking = SheetLocking()

  lazy val AllLocked: SheetLocking = SheetLocking(
    lockedAutoFilter = true,
    lockedDeleteColumns = true,
    lockedDeleteRows = true,
    lockedFormatCells = true,
    lockedFormatColumns = true,
    lockedFormatRows = true,
    lockedInsertColumns = true,
    lockedInsertHyperlinks = true,
    lockedInsertRows = true,
    lockedPivotTables = true,
    lockedSort = true,
    lockedObjects = true,
    lockedScenarios = true,
    lockedSelectLockedCells = true,
    lockedSelectUnlockedCells = true
  )

}

case class SheetLocking(
    lockedAutoFilter: Boolean = false,
    lockedDeleteColumns: Boolean = false,
    lockedDeleteRows: Boolean = false,
    lockedFormatCells: Boolean = false,
    lockedFormatColumns: Boolean = false,
    lockedFormatRows: Boolean = false,
    lockedInsertColumns: Boolean = false,
    lockedInsertHyperlinks: Boolean = false,
    lockedInsertRows: Boolean = false,
    lockedPivotTables: Boolean = false,
    lockedSort: Boolean = false,
    lockedObjects: Boolean = false,
    lockedScenarios: Boolean = false,
    lockedSelectLockedCells: Boolean = false,
    lockedSelectUnlockedCells: Boolean = false
) {

  def lockAutoFilter: SheetLocking = copy(lockedAutoFilter = true)

  def unlockAutoFilter: SheetLocking = copy(lockedAutoFilter = false)

  def lockDeleteColumns: SheetLocking = copy(lockedDeleteColumns = true)

  def unlockDeleteColumns: SheetLocking = copy(lockedDeleteColumns = false)

  def lockDeleteRows: SheetLocking = copy(lockedDeleteRows = true)

  def unlockDeleteRows: SheetLocking = copy(lockedDeleteRows = false)

  def lockFormatCells: SheetLocking = copy(lockedFormatCells = true)

  def unlockFormatCells: SheetLocking = copy(lockedFormatCells = false)

  def lockFormatColumns: SheetLocking = copy(lockedFormatColumns = true)

  def unlockFormatColumns: SheetLocking = copy(lockedFormatColumns = false)

  def lockFormatRows: SheetLocking = copy(lockedFormatRows = true)

  def unlockFormatRows: SheetLocking = copy(lockedFormatRows = false)

  def lockInsertColumns: SheetLocking = copy(lockedInsertColumns = true)

  def unlockInsertColumns: SheetLocking = copy(lockedInsertColumns = false)

  def lockInsertHyperlinks: SheetLocking = copy(lockedInsertHyperlinks = true)

  def unlockInsertHyperlinks: SheetLocking = copy(lockedInsertHyperlinks = false)

  def lockInsertRows: SheetLocking = copy(lockedInsertRows = true)

  def unlockInsertRows: SheetLocking = copy(lockedInsertRows = false)

  def lockPivotTables: SheetLocking = copy(lockedPivotTables = true)

  def unlockPivotTables: SheetLocking = copy(lockedPivotTables = false)

  def lockSort: SheetLocking = copy(lockedSort = true)

  def unlockSort: SheetLocking = copy(lockedSort = false)

  def lockObjects: SheetLocking = copy(lockedObjects = true)

  def unlockObjects: SheetLocking = copy(lockedObjects = false)

  def lockScenarios: SheetLocking = copy(lockedScenarios = true)

  def unlockScenarios: SheetLocking = copy(lockedScenarios = false)

  def lockSelectLockedCells: SheetLocking = copy(lockedSelectLockedCells = true)

  def unlockSelectLockedCells: SheetLocking = copy(lockedSelectLockedCells = false)

  def lockSelectUnlockedCells: SheetLocking = copy(lockedSelectUnlockedCells = true)

  def unlockSelectUnlockedCells: SheetLocking = copy(lockedSelectUnlockedCells = false)

}
