Data2XML is a new data capture tool which could be used to capture 
OMM data from any service that support RDM(Reuters Domain Models).
It could be used similar to Tick2XML, but there are only a little differences:

1. add new parameter "<data_type>" to indicate domain model, it should 
be one of the following values: "mp, mbp, mbo, mm, sl". ("mp" is default value 
if you do not set this parameter, currently do not support "sl". )

2. add new option "trace <file>" to log application events, mostly used 
check application running issues. (optional)

3. add new option "local_dict" to indicate whether should use local dictionaries 
to extract data. (optional, default to use dictionaries from service)

4. do not support following parameters from Tick2XML: "ordered, 
show_temp_num, no_signs, sslcfg, no_rst_fid_cache, no_resync_details".

====================================================

change history - v1.0.0.82

1. keep data2xml running and tip error if RMDS doesn't contain specified RDM.

2. do not filter update message in which the order size is 0 for MarketByOrder.

3. support Symbo List domain.

4. change output xml format of level2 snapshot message, include all items of one ric in a single snapshot.

5. change output xml format of item, use xml node for KEY and ACTION rather than using xml attribute.

====================================================

change history - v 1.0.0.88

1. Support getting configuration file and local dictionary from application root path.

====================================================

change history - v 1.0.0.89

1. Fix defect: Application will crash if its running root path is too long.

2. Support handle partial update of RMTES_STRING type for Market Price Domain.

====================================================

change history - v 1.0.0.90

1. Build based on RFA API v7.4.1.1 to reslove issues caused by RFA API.

2. Update RDMFieldDictionary/enumtype.def to TREP-RT 4.20.03.

