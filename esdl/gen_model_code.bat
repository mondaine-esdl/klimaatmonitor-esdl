REM C:\Users\matthijssenef\git\busdl\busdl\model\busdl.ecore

REM Ik denk dat het handiger is om -o . te gebruiken, zodat de imports van de gegenereerde files goed gaan
REM --auto-register-package will automatically register the metamodels when you import esdl or busdl
REM then no need to do 'rset.metamodel_registry[esdl.nsURI] = esdl' anymore
pyecoregen -vv -e model/esdl.ecore -o .
REM --auto-register-package