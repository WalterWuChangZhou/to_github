from write_forex import *
from write_XOP_NAV import *
from write_share_to_qdii_xlsx import *
from write_nav_to_qdII_xlsx import *
from write_close_amount_to_qdii_xlsx import *
from write_xop_price import *
from cal_add_share import *

update_usd_cny_safe()
xop_nav()
process_fund_close_amount_data()
write_lof_fund_nav()
process_fund_share_data()
update_xop_price()
update_add_share_xlsm()