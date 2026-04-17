defmodule ExVEx.MixProject do
  use Mix.Project

  @version "0.1.0"
  @source_url "https://github.com/jaman/ex_v_ex"

  def project do
    [
      app: :ex_v_ex,
      version: @version,
      elixir: "~> 1.18",
      start_permanent: Mix.env() == :prod,
      deps: deps(),
      description: description(),
      package: package(),
      docs: docs(),
      source_url: @source_url,
      name: "ExVEx",
      elixirc_paths: elixirc_paths(Mix.env())
    ]
  end

  def application do
    [extra_applications: [:logger]]
  end

  defp deps do
    [
      {:saxy, "~> 1.5"},
      {:ex_doc, "~> 0.34", only: :dev, runtime: false},
      {:dialyxir, "~> 1.4", only: [:dev, :test], runtime: false},
      {:credo, "~> 1.7", only: [:dev, :test], runtime: false}
    ]
  end

  defp description do
    "Pure Elixir library for reading and editing existing .xlsx / .xlsm files " <>
      "with round-trip fidelity — no Rust, no Python, no NIFs."
  end

  defp package do
    [
      licenses: ["MIT"],
      links: %{"GitHub" => @source_url},
      files: ~w(lib mix.exs README.md LICENSE CHANGELOG.md),
      maintainers: ["Jarius Jenkins"]
    ]
  end

  defp docs do
    [
      main: "ExVEx",
      source_ref: "v#{@version}",
      source_url: @source_url,
      extras: [
        "README.md",
        "CHANGELOG.md",
        LICENSE: [filename: "license", title: "License"]
      ]
    ]
  end

  defp elixirc_paths(:test), do: ["lib", "test/support"]
  defp elixirc_paths(_), do: ["lib"]
end
