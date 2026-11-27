/*
 * Sample data for each Style theme subject. Every sample has at least four
 * stages and balances (each box passes on exactly what it receives), and its
 * names are chosen to show off that theme's colour groups.
 */

export interface Sample {
  id: string;
  name: string;
  /** Value of the Style "Colour theme" dropdown to switch to. */
  theme: string;
  /** Header row first, then Source, Target, Value rows. */
  rows: (string | number)[][];
}

const HEADER = ["Source", "Target", "Value"];

export const SAMPLES: Sample[] = [
  {
    // A family's monthly budget (4 stages), in £: income into a current
    // account, split across household, spending, savings and card accounts,
    // then out to bills, everyday spending, savings pots and what's left.
    id: "cashflow",
    name: "Cashflow",
    theme: "theme:default",
    rows: [
      HEADER,
      ["Salary", "Current Account", 3800],
      ["Partner's Salary", "Current Account", 2600],
      ["Side Hustle", "Current Account", 450],
      ["Current Account", "Household Account", 2900],
      ["Current Account", "Spending Account", 1500],
      ["Current Account", "Savings Account", 1200],
      ["Current Account", "Credit Card", 800],
      ["Current Account", "Current Remainder", 450],
      ["Household Account", "Mortgage", 1450],
      ["Household Account", "Childcare", 830],
      ["Household Account", "Energy Bills", 240],
      ["Household Account", "Council Tax", 210],
      ["Household Account", "Car Insurance", 75],
      ["Household Account", "Water Bills", 55],
      ["Household Account", "Broadband", 40],
      ["Spending Account", "Groceries", 720],
      ["Spending Account", "Eating Out", 220],
      ["Spending Account", "Fuel", 180],
      ["Spending Account", "Kids' Clubs", 130],
      ["Spending Account", "Spending Remainder", 250],
      ["Savings Account", "Stocks & Shares ISA", 500],
      ["Savings Account", "Emergency Fund", 300],
      ["Savings Account", "Holiday Fund", 250],
      ["Savings Account", "Pension Top-up", 150],
      ["Credit Card", "Online Shopping", 290],
      ["Credit Card", "Travel", 150],
      ["Credit Card", "Gifts", 120],
      ["Credit Card", "Subscriptions", 60],
      ["Credit Card", "Card Buffer", 180],
    ],
  },
  {
    // An income statement (6 stages), in £k.
    id: "finance",
    name: "Finance",
    theme: "industry:finance",
    rows: [
      HEADER,
      ["Products", "Revenue", 300],
      ["Services", "Revenue", 120],
      ["Licensing", "Revenue", 80],
      ["Revenue", "Gross profit", 280],
      ["Revenue", "Cost of revenue", 220],
      ["Gross profit", "Operating profit", 160],
      ["Gross profit", "Operating expenses", 120],
      ["Operating expenses", "R&D", 50],
      ["Operating expenses", "Sales & marketing", 45],
      ["Operating expenses", "General & admin", 25],
      ["Operating profit", "Net profit", 118],
      ["Operating profit", "Income tax", 30],
      ["Operating profit", "Interest", 12],
      ["Net profit", "Dividends", 40],
      ["Net profit", "Retained earnings", 78],
    ],
  },
  {
    // Units through a distribution network (5 stages).
    id: "supply",
    name: "Supply chain",
    theme: "industry:supply",
    rows: [
      HEADER,
      ["Supplier A", "Inbound port", 400],
      ["Supplier B", "Inbound port", 250],
      ["Supplier C", "Central warehouse", 150],
      ["Inbound port", "Central warehouse", 630],
      ["Inbound port", "Damaged in transit", 20],
      ["Central warehouse", "North DC", 420],
      ["Central warehouse", "South DC", 330],
      ["Central warehouse", "Expired stock", 30],
      ["North DC", "Retail stores", 300],
      ["North DC", "Online orders", 110],
      ["North DC", "Customer returns", 10],
      ["South DC", "Retail stores", 250],
      ["South DC", "Online orders", 70],
      ["South DC", "Shrinkage", 10],
    ],
  },
  {
    // Kilograms of material through a production line (6 stages).
    id: "manufacturing",
    name: "Manufacturing",
    theme: "industry:manufacturing",
    rows: [
      HEADER,
      ["Steel coil", "Cutting", 600],
      ["Aluminium", "Cutting", 200],
      ["Plastic resin", "Moulding", 300],
      ["Cutting", "Welding", 720],
      ["Cutting", "Offcuts", 80],
      ["Moulding", "Assembly", 280],
      ["Moulding", "Rejected parts", 20],
      ["Welding", "Assembly", 690],
      ["Welding", "Rework", 30],
      ["Assembly", "Painting", 940],
      ["Assembly", "Defective units", 30],
      ["Painting", "Finished goods", 920],
      ["Painting", "Paint scrap", 20],
    ],
  },
  {
    // A recruitment funnel (6 stages), in candidates.
    id: "hr",
    name: "HR",
    theme: "industry:hr",
    rows: [
      HEADER,
      ["LinkedIn", "Applications", 320],
      ["Job boards", "Applications", 260],
      ["Referrals", "Applications", 90],
      ["Agency", "Applications", 80],
      ["Applications", "Phone screen", 240],
      ["Applications", "Rejected at CV", 480],
      ["Applications", "Withdrawn", 30],
      ["Phone screen", "Interview", 110],
      ["Phone screen", "Not progressed", 120],
      ["Phone screen", "Withdrawn", 10],
      ["Interview", "Offer", 30],
      ["Interview", "Rejected after interview", 75],
      ["Interview", "Withdrawn", 5],
      ["Offer", "Hired", 24],
      ["Offer", "Offer declined", 6],
    ],
  },
  {
    // A sales pipeline (6 stages), in leads.
    id: "sales",
    name: "Sales",
    theme: "industry:sales",
    rows: [
      HEADER,
      ["Website", "Leads", 600],
      ["Paid ads", "Leads", 400],
      ["Events", "Leads", 150],
      ["Partner referrals", "Leads", 100],
      ["Leads", "Qualified", 500],
      ["Leads", "Disqualified", 600],
      ["Leads", "Nurture", 150],
      ["Qualified", "Demo", 300],
      ["Qualified", "No response", 200],
      ["Demo", "Proposal", 180],
      ["Demo", "Lost", 120],
      ["Proposal", "Won", 80],
      ["Proposal", "Lost", 70],
      ["Proposal", "On hold", 30],
    ],
  },
];

export const sampleById = (id: string) => SAMPLES.find((s) => s.id === id);
