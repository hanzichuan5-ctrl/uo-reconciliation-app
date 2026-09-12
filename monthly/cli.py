"""The UI and command line use the same monthly processing functions."""
import argparse
import json
from .core import bootstrap,build_month
from .exporter import export_month

def main():
    parser=argparse.ArgumentParser(description='检查或导出已导入资料的月度结果')
    parser.add_argument('--month',required=True)
    parser.add_argument('--export',choices=['income','receipt','both'])
    args=parser.parse_args()
    bootstrap();data=build_month(args.month)
    result={'month':args.month,'channelInvoiceCount':data['income']['inScopeCount'],
            'scopeReviewCount':data['income']['scopeReviewCount'],
            'includedReceipts':data['receipts']['stats'].get('included',0),'receiptAmountCents':data['receipts']['amountCents']}
    if args.export:
        result['exports']=[]
        for kind in (['income','receipt'] if args.export=='both' else [args.export]):
            try:
                result['exports'].append(export_month(args.month,kind))
            except ValueError as exc:
                result['exports'].append({'kind':kind,'status':'missing_input','reason':str(exc)})
    print(json.dumps(result,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
