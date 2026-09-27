# Publishing the model (GitHub + Zenodo DOI)

JDR asks for the model to be deposited in a public git repository. A DOI-backed archive of the
exact release cited in the paper is the standard way to do this.

## 1. Create the public repository

```bash
cd adolescent-sealant-cea
git init
git add .
git commit -m "Model behind 'Impact and Cost-Effectiveness of Adolescent School Sealant Programs' (submitted version + PSA)"
# on github.com: New repository → name: adolescent-sealant-cea → Public → create (no README)
git remote add origin https://github.com/<org>/adolescent-sealant-cea.git
git branch -M main
git push -u origin main
```

Then replace `<org>` in `README.md` and `CITATION.cff` with the actual owner (a lab/organisation
account is preferable to a personal one for longevity), commit and push again.

## 2. Tag the version cited in the manuscript

```bash
git tag -a v1.0.0 -m "Version used for the JDR revision"
git push origin v1.0.0
```

## 3. Mint a DOI with Zenodo

1. Sign in at https://zenodo.org with GitHub, open *GitHub* under your account settings and switch the
   repository **on**.
2. On GitHub, *Releases → Draft a new release → choose tag v1.0.0 → Publish release*.
   Zenodo archives the release automatically and assigns a DOI (plus a concept DOI that always resolves
   to the latest version).
3. Paste the DOI into `CITATION.cff` (`doi:` line) and into the manuscript's data/code-availability
   statement; commit, push, and (optionally) publish a v1.0.1 release so the archive itself carries the DOI.

## 4. Manuscript statement (suggested wording)

> **Code availability.** The Python model, all input parameters, and the scripts that regenerate every
> table and figure in this article (including the probabilistic sensitivity analysis) are openly
> available under the MIT license at https://github.com/<org>/adolescent-sealant-cea and are archived at
> Zenodo (https://doi.org/10.5281/zenodo.XXXXXXX).

## 5. Before pushing — checklist

- [ ] Author list and affiliations in `CITATION.cff` and `LICENSE`
- [ ] `<org>` replaced everywhere (`grep -rn "<org>" .`)
- [ ] `IMPL_SCREENING_COSTS` in `model/psa_ceac.py` set to the fee-schedule value(s) you decide to report
- [ ] `python model/sealant_markov_model.py` prints `62/62` and `26/26`
- [ ] `results/` regenerated from a clean checkout (`git clean -xdf results && python model/...`)
