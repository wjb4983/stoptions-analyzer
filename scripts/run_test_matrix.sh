#!/usr/bin/env bash
set -u

REPORT_ROOT="reports/test_matrix"
ENABLE_COVERAGE="${TEST_MATRIX_COVERAGE:-0}"
PYTEST_BIN="${PYTEST_BIN:-pytest}"

mkdir -p "${REPORT_ROOT}"

TIERS=(
  "smoke|smoke|--maxfail=1"
  "core|core|"
  "core_or_slow|core or slow|"
)

printf "Running test matrix with %s tiers...\n" "${#TIERS[@]}"
printf "Artifacts: %s\n\n" "${REPORT_ROOT}"

EXIT_CODE=0
SUMMARY_ROWS=()

for tier in "${TIERS[@]}"; do
  IFS='|' read -r tier_name marker extra_flags <<< "${tier}"

  tier_dir="${REPORT_ROOT}/${tier_name}"
  mkdir -p "${tier_dir}"

  log_file="${tier_dir}/stdout.log"
  junit_file="${tier_dir}/junit.xml"

  coverage_flags=()
  if [[ "${ENABLE_COVERAGE}" == "1" ]]; then
    coverage_flags=(
      "--cov=."
      "--cov-report=xml:${tier_dir}/coverage.xml"
      "--cov-report=html:${tier_dir}/coverage_html"
    )
  fi

  cmd=(
    "${PYTEST_BIN}"
    -m "${marker}"
    -ra
    --junitxml "${junit_file}"
  )

  if [[ -n "${extra_flags}" ]]; then
    # shellcheck disable=SC2206
    extra_arr=(${extra_flags})
    cmd+=("${extra_arr[@]}")
  fi

  cmd+=("${coverage_flags[@]}")

  start_ts="$(date +%s)"
  printf "[%s] %s\n" "${tier_name}" "${cmd[*]}" | tee "${log_file}"

  if "${cmd[@]}" >> "${log_file}" 2>&1; then
    status="PASS"
  else
    status="FAIL"
    EXIT_CODE=1
  fi

  end_ts="$(date +%s)"
  duration="$(( end_ts - start_ts ))"

  SUMMARY_ROWS+=("${tier_name}|${status}|${duration}s|${log_file}|${junit_file}")

  printf "[%s] status=%s duration=%ss\n\n" "${tier_name}" "${status}" "${duration}" | tee -a "${log_file}"
done

printf "\nTest matrix summary\n"
printf "%-14s %-8s %-10s %-40s\n" "Tier" "Status" "Duration" "Artifacts"
printf "%-14s %-8s %-10s %-40s\n" "--------------" "--------" "----------" "----------------------------------------"
for row in "${SUMMARY_ROWS[@]}"; do
  IFS='|' read -r tier_name status duration log_file junit_file <<< "${row}"
  printf "%-14s %-8s %-10s %s, %s\n" "${tier_name}" "${status}" "${duration}" "${log_file}" "${junit_file}"
done

exit "${EXIT_CODE}"
