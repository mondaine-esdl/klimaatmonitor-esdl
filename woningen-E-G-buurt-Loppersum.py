from xlrd import open_workbook, cellname
from pyecore.resources import ResourceSet, URI
from esdl.esdl import esdl
from xmlresource import XMLResource
from stringuri import StringURI
import uuid
import sys
import re

gas_to_heat_efficiency = 0.9
kwh_to_tj = 3.6e-6
gas_energy_content = 35.17 # MJ/m3
mega_to_terra = 1e6
gas_heater_efficiency = 0.9
gas_heater_power = 10000000.0
elec_prod_mc = 0.9
elec_cons_mc = 0.1
gas_prod_mc = 0.9
gas_cons_mc = 0.1
gep_power = 1e12
gec_power = 1e12
ggp_power = 1e12
ggc_power = 1e12

profileType = "ENERGY_IN_TJ"
influx_host = "http://10.30.2.1"
influx_port = 8086
influx_database = "energy_profiles"
influx_filters = ""
elec_measurement = "nedu_elektriciteit_2015-2018"
elec_field = "E1A"
elec_field_comp = "E3A"     # Assume offices, shops, education
gas_measurement = "nedu_aardgas_2015-2018"
gas_field = "G1A"
gas_field_comp = "G2A"


def ci2s(cell):
    return str(int(cell.value))


def excel_to_ESDL(fname, sname_woningen, sname_bedrijven):
    book = open_workbook(fname)
    sheet = book.sheet_by_name(sname_woningen)

    top_area_type = str(sheet.cell(0, 3).value)
    if top_area_type == 'GM':
        top_area_scope = 'MUNICIPALITY'
    else:
        print('Other scopes than municipality not supported yet')
        # sys.exit(1)

    top_area_code = top_area_type + str(sheet.cell(0, 4).value)
    top_area_name = str(sheet.cell(0, 1).value)
    top_area_year = ci2s(sheet.cell(0, 2))
    sub_aggr = str(sheet.cell(0, 5).value)
    if sub_aggr == 'WK':
        sub_aggr_id_start = 'WK'+str(sheet.cell(0, 4).value)
        sub_aggr_scope = 'DISTRICT'
        sub_area_number_pos = 5
    elif sub_aggr[0:2] == 'BU':
        sub_aggr_scope = 'NEIGHBOURHOOD'
        sub_aggr_code_column = 3
        print('TODO: check buurt codes in excel')

    column1_name = str(sheet.cell(1, 1).value)
    column2_name = str(sheet.cell(1, 2).value)

    if column1_name[0:3] == 'gas':
        gas_column = 1
    if column1_name[0:4] == 'elek':
        elec_column = 1
    if column2_name[0:3] == 'gas':
        gas_column = 2
    if column2_name[0:4] == 'elek':
        elec_column = 2

    es = esdl.EnergySystem(id = str(uuid.uuid4()), name = top_area_name + ' ' + top_area_year)

    carrs = esdl.Carriers(id=str(uuid.uuid4()))

    elec_car = esdl.EnergyCarrier(id='ELEC', name='Electricity', emission=180.28, energyContent=1.0, energyCarrierType='FOSSIL')
    elec_car.emissionUnit = esdl.QuantityAndUnitType(physicalQuantity='EMISSION', multiplier='KILO', unit='GRAM', perMultiplier='GIGA', perUnit='JOULE')
    elec_car.energyContentUnit = esdl.QuantityAndUnitType(physicalQuantity='ENERGY', multiplier='MEGA', unit='JOULE', perMultiplier='MEGA', perUnit='JOULE')
    carrs.carrier.append(elec_car)

    gas_car = esdl.EnergyCarrier(id='GAS', name='Natural Gas', emission=1.788225, energyContent=35.1700000, energyCarrierType='FOSSIL')
    gas_car.emissionUnit = esdl.QuantityAndUnitType(physicalQuantity='EMISSION', multiplier='KILO', unit='GRAM', perUnit='CUBIC_METRE')
    gas_car.energyContentUnit = esdl.QuantityAndUnitType(physicalQuantity='ENERGY', multiplier='MEGA', unit='JOULE', perUnit='CUBIC_METRE')
    carrs.carrier.append(gas_car)

    heat_comm = esdl.HeatCommodity(id='HEAT', name='Heat')
    carrs.carrier.append(heat_comm)

    es.energySystemInformation = esdl.EnergySystemInformation(id=str(uuid.uuid4()), carriers=carrs)

    srvs = esdl.Services()

    inst = esdl.Instance(id = str(uuid.uuid4()), name = top_area_name + ' ' + top_area_year)
    es.instance.append(inst)
    ar = esdl.Area(id = top_area_code, name = top_area_name, scope = top_area_scope)
    es.instance[0].area = ar

    elec_nw_op = esdl.OutPort(id=str(uuid.uuid4()), name='EOut', carrier=elec_car)
    elec_nw_ip = esdl.InPort(id=str(uuid.uuid4()), name='EIn', carrier=elec_car)
    elec_nw = esdl.ElectricityNetwork(id=str(uuid.uuid4()), name='ElectricityNetwork', port=[elec_nw_ip, elec_nw_op])
    ar.asset.append(elec_nw)

    gep_mc = esdl.SingleValue(id=str(uuid.uuid4()), name='MarginalCosts', value=elec_prod_mc)
    gep_ci = esdl.CostInformation(marginalCosts=gep_mc)
    gep_op = esdl.OutPort(id=str(uuid.uuid4()), name='EOut', connectedTo=[elec_nw_ip], carrier=elec_car)
    gep = esdl.GenericProducer(id=str(uuid.uuid4()), name='Unlimited Electricity Generation', port=[gep_op],
                               power=gep_power, costInformation=gep_ci, prodType=esdl.RenewableTypeEnum.FOSSIL)
    ar.asset.append(gep)
    gec_mc = esdl.SingleValue(id=str(uuid.uuid4()), name='MarginalCosts', value=elec_cons_mc)
    gec_ci = esdl.CostInformation(marginalCosts=gec_mc)
    gec_ip = esdl.InPort(id=str(uuid.uuid4()), name='EIn', connectedTo=[elec_nw_op], carrier=elec_car)
    gec = esdl.GenericConsumer(id=str(uuid.uuid4()), name='Unlimited Electricity Consumption', port=[gec_ip],
                               power=gec_power, costInformation=gec_ci)
    ar.asset.append(gec)

    gas_nw_op = esdl.OutPort(id=str(uuid.uuid4()), name='GOut', carrier=gas_car)
    gas_nw_ip = esdl.InPort(id=str(uuid.uuid4()), name='GIn', carrier=gas_car)
    gas_nw = esdl.GasNetwork(id=str(uuid.uuid4()), name='GasNetwork', port=[gas_nw_ip, gas_nw_op])
    ar.asset.append(gas_nw)

    ggp_mc = esdl.SingleValue(id=str(uuid.uuid4()), name='MarginalCosts', value=gas_prod_mc)
    ggp_ci = esdl.CostInformation(marginalCosts=ggp_mc)
    ggp_op = esdl.OutPort(id=str(uuid.uuid4()), name='GOut', connectedTo=[gas_nw_ip], carrier=gas_car)
    ggp = esdl.GenericProducer(id=str(uuid.uuid4()), name='Unlimited Gas Generation', port=[ggp_op],
                               power=ggp_power, costInformation=ggp_ci, prodType=esdl.RenewableTypeEnum.FOSSIL)
    ar.asset.append(ggp)
    # ggc_mc = esdl.SingleValue(id=str(uuid.uuid4()), name='MarginalCosts', value=gas_cons_mc)
    # ggc_ci = esdl.CostInformation(marginalCosts=ggc_mc)
    # ggc_ip = esdl.InPort(id=str(uuid.uuid4()), name='EIn', connectedTo=[gas_nw_op], carrier=elec_car)
    # ggc = esdl.GenericConsumer(id=str(uuid.uuid4()), name='Unlimited Gas Consumption', port=[ggc_ip],
    #                            power=ggc_power, costInformation=ggc_ci)
    # ar.asset.append(ggc)

    for row in range(2, sheet.nrows-3):
        sub_area_name = str(sheet.cell(row, 0).value)
        if sub_aggr_scope == 'DISTRICT':
            sub_area_number = sub_area_name[sub_area_number_pos:sub_area_number_pos+2]
            sub_area_id = str(sub_aggr_id_start + sub_area_number)
        else:
            sub_area_id = str(sheet.cell(row, sub_aggr_code_column).value)

        gas_value = sheet.cell(row, gas_column).value
        if str(gas_value) != '?' and str(gas_value) != '':
            heat_value = gas_heater_efficiency * gas_value
        else:
            heat_value = None
        elec_value = sheet.cell(row, elec_column).value
        if str(elec_value) == '?' and str(elec_value) != '':
            elec_value = None

        sub_area = esdl.Area(id=sub_area_id, name=sub_area_name, scope=sub_aggr_scope)

        aggr_build = esdl.AggregatedBuilding(id=str(uuid.uuid4()), name="building")

        if heat_value:
            hdprofqau = esdl.QuantityAndUnitType(physicalQuantity='ENERGY', multiplier='TERRA', unit='JOULE')
            hdprof = esdl.InfluxDBProfile(id=str(uuid.uuid4()), multiplier=heat_value, host=influx_host,
                                          port=influx_port, database=influx_database, filters=influx_filters,
                                          measurement=gas_measurement, field=gas_field, profileQuantityAndUnit=hdprofqau,
                                          profileType='ENERGY_IN_TJ')
            hdip = esdl.InPort(id=str(uuid.uuid4()), name='InPort', carrier=heat_comm, profile=hdprof)
            hd = esdl.HeatingDemand(id=str(uuid.uuid4()), name='HeatingDemand_'+sub_area_id, port=[hdip])

            ghip = esdl.InPort(id=str(uuid.uuid4()), name='InPort', connectedTo=[gas_nw_op], carrier=gas_car)
            ghop = esdl.OutPort(id=str(uuid.uuid4()), name='OutPort', connectedTo=[hdip], carrier=heat_comm)
            gh = esdl.GasHeater(id=str(uuid.uuid4()), name='GasHeater_'+sub_area_id, efficiency=gas_heater_efficiency,
                                power=gas_heater_power, port=[ghip, ghop])

            dbd = esdl.DrivenByDemand(id=str(uuid.uuid4()), name='DBD_GasHeater_'+sub_area_id, energyAsset=gh, outPort=ghop)
            srvs.service.append(dbd)
            aggr_build.asset.append(hd)
            aggr_build.asset.append(gh)

        if elec_value:
            edprofqau = esdl.QuantityAndUnitType(physicalQuantity='ENERGY', multiplier='TERRA', unit='JOULE')
            edprof = esdl.InfluxDBProfile(id=str(uuid.uuid4()), multiplier=elec_value, host=influx_host,
                                          port=influx_port, database=influx_database, filters=influx_filters,
                                          measurement=elec_measurement, field=elec_field, profileQuantityAndUnit=edprofqau,
                                          profileType='ENERGY_IN_TJ')
            edip = esdl.InPort(id=str(uuid.uuid4()), name='InPort', connectedTo=[elec_nw_op], carrier=elec_car, profile=edprof)
            ed = esdl.ElectricityDemand(id=str(uuid.uuid4()), name='ElectricityDemand_'+sub_area_id, port=[edip])

            aggr_build.asset.append(ed)

        if heat_value or elec_value:
            sub_area.asset.append(aggr_build)

        ar.area.append(sub_area)

    if sname_bedrijven:
        sheet = book.sheet_by_name(sname_bedrijven)
        aggr_build_comp = esdl.AggregatedBuilding(id=str(uuid.uuid4()), name="building-bedrijven")

        for row in range(2, sheet.nrows-3):
            cat = str(sheet.cell(row, 0).value)
            waarde = sheet.cell(row, 1).value

            if cat != '' and waarde != 0 and waarde != '?':     # filter non relevant data
                cat = re.sub(' ', '_', cat)
                print(cat, waarde)

                if cat.find('m3') != -1:        # category contains gasusage
                    gas_tj = waarde * gas_energy_content / mega_to_terra
                    heat_tj = gas_tj * gas_to_heat_efficiency

                    hdbprofqau = esdl.QuantityAndUnitType(physicalQuantity='ENERGY', multiplier='TERRA', unit='JOULE')
                    hdbprof = esdl.InfluxDBProfile(id=str(uuid.uuid4()), multiplier=heat_tj, host=influx_host,
                                                   port=influx_port, database=influx_database, filters=influx_filters,
                                                   measurement=gas_measurement, field=gas_field_comp,
                                                   profileQuantityAndUnit=hdbprofqau,
                                                   profileType='ENERGY_IN_TJ')
                    hdbip = esdl.InPort(id=str(uuid.uuid4()), name='InPort', carrier=heat_comm, profile=hdbprof)
                    hdb = esdl.HeatingDemand(id=str(uuid.uuid4()), name='HeatingDemand_' + cat, port=[hdbip])

                    ghbip = esdl.InPort(id=str(uuid.uuid4()), name='InPort', connectedTo=[gas_nw_op], carrier=gas_car)
                    ghbop = esdl.OutPort(id=str(uuid.uuid4()), name='OutPort', connectedTo=[hdbip], carrier=heat_comm)
                    ghb = esdl.GasHeater(id=str(uuid.uuid4()), name='GasHeater_' + cat, efficiency=gas_heater_efficiency,
                                         power=gas_heater_power, port=[ghbip, ghbop])

                    aggr_build_comp.asset.append(hdb)
                    aggr_build_comp.asset.append(ghb)

                    dbd = esdl.DrivenByDemand(id=str(uuid.uuid4()), name='DBD_GasHeater_' + cat, energyAsset=ghb,
                                              outPort=ghbop)
                    srvs.service.append(dbd)

                if cat.find('kWh') != -1:       # category contains electricity usage
                    elec_tj = waarde * kwh_to_tj

                    edbprofqau = esdl.QuantityAndUnitType(physicalQuantity='ENERGY', multiplier='TERRA', unit='JOULE')
                    edbprof = esdl.InfluxDBProfile(id=str(uuid.uuid4()), multiplier=elec_tj, host=influx_host,
                                                   port=influx_port, database=influx_database, filters=influx_filters,
                                                   measurement=elec_measurement, field=elec_field_comp, profileQuantityAndUnit=edbprofqau,
                                                   profileType='ENERGY_IN_TJ')
                    edbip = esdl.InPort(id=str(uuid.uuid4()), name='InPort', connectedTo=[elec_nw_op], carrier=elec_car, profile=edbprof)
                    edb = esdl.ElectricityDemand(id=str(uuid.uuid4()), name='ElectricityDemand_'+cat, port=[edbip])

                    aggr_build_comp.asset.append(edb)

        ar.asset.append(aggr_build_comp)

    es.services = srvs

    rset = ResourceSet()
    rset.resource_factory['esdl'] = lambda uri: XMLResource(uri)
    rset.metamodel_registry[esdl.nsURI] = esdl
    resource = rset.create_resource(URI(top_area_name+top_area_year+'.esdl'))
    resource.append(es)
    resource.save()


def main():

    fname = './data/Totaal  2017 - Buurten 2018 van Loppersum.xls'
    sname_won = 'Totaal 2017 Buurten 2018 van Lo'

    excel_to_ESDL(fname, sname_won, None)


if __name__ == "__main__":
    main()