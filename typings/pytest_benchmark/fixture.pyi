from collections.abc import Callable

class BenchmarkFixture:
    def __call__[**P, R](
        self,
        function_to_benchmark: Callable[P, R],
        *args: P.args,
        **kwargs: P.kwargs,
    ) -> R: ...
