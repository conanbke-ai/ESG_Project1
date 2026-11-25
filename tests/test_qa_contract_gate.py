import unittest
from tools.qa_contract_gate import problems

def valid():
    return {
        "feature":"sample-change", "project":"nongtori", "risk":"MEDIUM",
        "user_facing":True, "acceptance_criteria":["AC-1"],
        "qa_roles":["QA Lead & Test Architect","Functional & Regression QA Reviewer",
                    "Product Flow QA Reviewer","Acceptance Reviewer"],
        "test_cases":["positive","negative"], "regression_scope":"adjacent API",
        "evidence":[{"check":"unit tests","state":"PASS","proof":"CI job URL / command output"}],
        "not_run":[], "exit_decision":"PASS"
    }

class ContractTests(unittest.TestCase):
    def test_valid(self):
        self.assertEqual(problems(valid()), [])
    def test_missing_precheck(self):
        d=valid();d.pop("acceptance_criteria")
        self.assertTrue(problems(d))
    def test_review_is_not_test(self):
        d=valid();d["evidence"]=[{"check":"static review","state":"REVIEWED"}]
        self.assertTrue(problems(d))
    def test_not_run_blocks(self):
        d=valid();d["not_run"]=["browser acceptance"]
        self.assertTrue(problems(d))
    def test_missing_flow_role_blocks(self):
        d=valid();d["qa_roles"].remove("Product Flow QA Reviewer")
        self.assertTrue(problems(d))
    def test_partial_blocks(self):
        d=valid();d["exit_decision"]="PARTIAL"
        self.assertTrue(problems(d))

if __name__ == "__main__":
    unittest.main()
