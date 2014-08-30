TCLSH = tclsh
RM    = del
RMDIR = rmdir /S /Q
CP    = copy
CPDIR = xcopy /E /I /Y /Q
MKDIR = md
DSEP  = \\

MYDIR=$(subst /,\\,$(CURDIR))
DISTDIR=$(MYDIR)$(DSEP)..$(DSEP)CawtDistribution

# Installation directory usable for CAWT developers.
# Adapt to your local needs.
INSTDIR=C:\opt\poSoft\lib\Cawt

dist: kit doc distcawt

clean:
	-$(RMDIR) TestPrograms$(DSEP)testOut
	-$(RM) TestPrograms$(DSEP)RunTests.log

distclean: clean
	-cd Starkit && $(TCLSH) makeKit.tcl distclean win32
	-cd Starkit && $(TCLSH) makeKit.tcl distclean win64
	-cd Documentation && $(TCLSH) genCawtDoc.tcl distclean

test:
	cd TestPrograms && RunTests.bat

kit:
	cd Starkit && $(TCLSH) makeKit.tcl all win32
	cd Starkit && $(TCLSH) makeKit.tcl all win64

doc:
	cd Documentation && $(TCLSH) genCawtDoc.tcl

distcawt:
	$(TCLSH) makeDist.tcl $(DISTDIR)

install:
	$(MKDIR) $(INSTDIR)
	$(CPDIR) CawtCore     $(INSTDIR)$(DSEP)CawtCore
	$(CPDIR) CawtEarth    $(INSTDIR)$(DSEP)CawtEarth
	$(CPDIR) CawtExcel    $(INSTDIR)$(DSEP)CawtExcel
	$(CPDIR) CawtExplorer $(INSTDIR)$(DSEP)CawtExplorer
	$(CPDIR) CawtMatlab   $(INSTDIR)$(DSEP)CawtMatlab
	$(CPDIR) CawtOcr      $(INSTDIR)$(DSEP)CawtOcr
	$(CPDIR) CawtPpt      $(INSTDIR)$(DSEP)CawtPpt
	$(CPDIR) CawtOutlook  $(INSTDIR)$(DSEP)CawtOutlook
	$(CPDIR) CawtWord     $(INSTDIR)$(DSEP)CawtWord
	$(CP)    pkgIndex.tcl $(INSTDIR)$(DSEP)pkgIndex.tcl
