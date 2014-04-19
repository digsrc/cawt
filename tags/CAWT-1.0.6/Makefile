TCLSH = tclsh
RM    = del
RMDIR = rmdir /S /Q
CP    = copy
CPDIR = xcopy /E /I /Y /Q
MKDIR = md
DSEP  = \\

MYDIR=$(subst /,\\,$(CURDIR))
DISTDIR=$(MYDIR)$(DSEP)..$(DSEP)CawtDistribution

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
