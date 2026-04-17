defmodule ExVEx.Test.Fixtures do
  @moduledoc false

  @fixture_dir Path.expand("../fixtures", __DIR__)

  @spec path(String.t()) :: String.t()
  def path(name), do: Path.join(@fixture_dir, name)

  @spec tmp_path(String.t()) :: String.t()
  def tmp_path(name) do
    suffix = :erlang.unique_integer([:positive, :monotonic])
    Path.join(System.tmp_dir!(), "ex_v_ex_#{suffix}_#{name}")
  end
end
