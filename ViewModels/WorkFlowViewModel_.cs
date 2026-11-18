using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using WCals.Models;

namespace WCals.ViewModels
{
    public class WorkFlowViewModel
    {
        public IEnumerable<SMCCODE> Actions { get; set; }
        public IEnumerable<SMCCODE> Steps { get; set; }
        public IEnumerable<OGCDEPT> Depts { get; set; }
        public IEnumerable<DeptUsersViewModel> DeptUsers { get; set; }
        public IEnumerable<OGCUSER> Users { get; set; }
        public IEnumerable<TDV_FCWPOOL> FCWPOOL { get; set; }
        public string TD_CODE { get; set; }
        public string STEP_NO { get; set; }
        public string DEPT_NO { get; set; }
        public string SC_NO { get; set; }
        public string ACTION { get; set; }
        public string USR_ID { get; set; }
        public byte[] fileArray { get; set; }
        //Tender
        public IEnumerable<TDV_TDMFORM> TDV_TDMFORM { get; set; }
        public IEnumerable<TDMFORM> TDMFORM { get; set; }
        public IEnumerable<TDDNPAN> TDDNPAN { get; set; }
        public IEnumerable<TDDCQTY> TDDCQTY { get; set; }
        public IEnumerable<TDDECWT> TDDECWT { get; set; }
        public IEnumerable<TDDYPSC> TDDYPSC { get; set; }
        public IEnumerable<TDHBGLS> TDHBGLS { get; set; }
        public IEnumerable<TDDYPAC> TDDYPAC { get; set; }
        public IEnumerable<TDDMSCD> TDDMSCD { get; set; }
        public IEnumerable<TDDACCT> TDDACCT { get; set; }
        public IEnumerable<TDDACPT> TDDACPT { get; set; }
        public IEnumerable<TDDARBK> TDDARBK { get; set; }
        public IEnumerable<TDDARRC> TDDARRC { get; set; }
        public IEnumerable<TDDARRE> TDDARRE { get; set; }
        public IEnumerable<TDDBDCM> TDDBDCM { get; set; }
        public IEnumerable<TDDBDRC> TDDBDRC { get; set; }
        public IEnumerable<TDDBGET> TDDBGET { get; set; }
        public IEnumerable<TDDBGLS> TDDBGLS { get; set; }
        public IEnumerable<TDDBGRC> TDDBGRC { get; set; }
        public IEnumerable<TDDCHKR> TDDCHKR { get; set; }
        public IEnumerable<TDDCPGP> TDDCPGP { get; set; }
        public IEnumerable<TDDDRCR> TDDDRCR { get; set; }
        public IEnumerable<TDDECCR> TDDECCR { get; set; }
        public IEnumerable<TDDECRF> TDDECRF { get; set; }
        // public IEnumerable<TDDEXFD> TDDEXFD { get; set; }
        public IEnumerable<TDDFDIS> TDDFDIS { get; set; }
        public IEnumerable<TDDFDSC> TDDFDSC { get; set; }
        public IEnumerable<TDDFUND> TDDFUND { get; set; }
        public IEnumerable<TDDGCRC> TDDGCRC { get; set; }
        public IEnumerable<TDDHODY> TDDHODY { get; set; }
        public IEnumerable<TDDIOMT> TDDIOMT { get; set; }
        public IEnumerable<TDDMATL> TDDMATL { get; set; }
        public IEnumerable<TDDOB> TDDOB { get; set; }
        // public IEnumerable<TDDPER> TDDPER { get; set; }
        public IEnumerable<TDDPLDJ> TDDPLDJ { get; set; }
        public IEnumerable<TDDPRMT> TDDPRMT { get; set; }
        public IEnumerable<TDDREVF> TDDREVF { get; set; }
        public IEnumerable<TDDRMPR> TDDRMPR { get; set; }
        public IEnumerable<TDDSCHD> TDDSCHD { get; set; }
        public IEnumerable<TDDSTWR> TDDSTWR { get; set; }
        public IEnumerable<TDDSUPM> TDDSUPM { get; set; }
        public IEnumerable<TDDUCLF> TDDUCLF { get; set; }
        public IEnumerable<TDDWBSC> TDDWBSC { get; set; }
        public IEnumerable<TDDWKEN> TDDWKEN { get; set; }
        public IEnumerable<TDDWKLS> TDDWKLS { get; set; }
        public IEnumerable<TDDYPFD> TDDYPFD { get; set; }
        public IEnumerable<TDHBGET> TDHBGET { get; set; }
        public IEnumerable<TDHYPAC> TDHYPAC { get; set; }
        public IEnumerable<TDHYPFD> TDHYPFD { get; set; }
        public IEnumerable<TDHYPSC> TDHYPSC { get; set; }
        public IEnumerable<TDMIOMT> TDMIOMT { get; set; }
        public IEnumerable<TDMPRMT> TDMPRMT { get; set; }
        public IEnumerable<TDWDTSM> TDWDTSM { get; set; }
    }
}