"""DNS NS record lookup logic with concurrent resolution."""

import dns.resolver
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Callable


def lookup_nameservers(domain: str, timeout: float = 5.0) -> dict:
    """Look up NS records for a single domain."""
    domain = domain.strip().lower()
    if not domain:
        return {"domain": domain, "nameservers": [], "ns_count": 0, "status": "SKIPPED"}

    resolver = dns.resolver.Resolver()
    resolver.lifetime = timeout
    resolver.timeout = timeout

    try:
        answers = resolver.resolve(domain, "NS")
        nameservers = sorted(str(r.target).rstrip(".") for r in answers)
        return {
            "domain": domain,
            "nameservers": nameservers,
            "ns_count": len(nameservers),
            "status": "OK",
        }
    except dns.resolver.NXDOMAIN:
        return {"domain": domain, "nameservers": [], "ns_count": 0, "status": "NXDOMAIN"}
    except dns.resolver.NoAnswer:
        return {"domain": domain, "nameservers": [], "ns_count": 0, "status": "NO_NS_RECORD"}
    except dns.resolver.Timeout:
        return {"domain": domain, "nameservers": [], "ns_count": 0, "status": "TIMEOUT"}
    except dns.exception.DNSException as e:
        return {"domain": domain, "nameservers": [], "ns_count": 0, "status": f"ERROR: {e}"}


def run_bulk_lookup(
    domains: list[str],
    workers: int = 20,
    timeout: float = 5.0,
    progress_cb: Callable[[int, int], None] | None = None,
) -> list[dict]:
    """
    Concurrently look up NS records for a list of domains.
    Calls progress_cb(completed, total) after each result.
    Returns results in the same order as input domains.
    """
    total = len(domains)
    results = [None] * total

    with ThreadPoolExecutor(max_workers=workers) as executor:
        future_to_idx = {
            executor.submit(lookup_nameservers, d, timeout): i
            for i, d in enumerate(domains)
        }
        completed = 0
        for future in as_completed(future_to_idx):
            idx = future_to_idx[future]
            results[idx] = future.result()
            completed += 1
            if progress_cb:
                progress_cb(completed, total)

    return results
