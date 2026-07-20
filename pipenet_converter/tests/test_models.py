"""Tests for core network data models."""

import pytest

from pipenet_converter.models import (
    Equipment,
    Fitting,
    Node,
    Nozzle,
    Pipe,
    PipeNetwork,
    Valve,
    diameter_label_to_m,
    diameter_m_to_label,
)


def test_diameter_label_to_m_converts_supported_labels() -> None:
    assert diameter_label_to_m("25A") == 0.025
    assert diameter_label_to_m("150A") == 0.15
    assert diameter_label_to_m("200A") == 0.2


def test_diameter_m_to_label_converts_supported_diameters() -> None:
    assert diameter_m_to_label(0.032) == "32A"
    assert diameter_m_to_label(0.15) == "150A"


def test_unsupported_diameter_values_raise_value_error() -> None:
    with pytest.raises(ValueError, match="Unsupported diameter label"):
        diameter_label_to_m("300A")

    with pytest.raises(ValueError, match="Unsupported diameter"):
        diameter_m_to_label(0.3)


def test_pipe_network_adds_and_retrieves_objects() -> None:
    network = PipeNetwork(title="Sample network")
    node_1 = Node("N001", 0.0, 0.0, 48.55, "pipe")
    node_2 = Node("N002", 10.0, 0.0, 48.55, "head")
    pipe = Pipe(
        pipe_id="P001",
        from_node="N001",
        to_node="N002",
        diameter_m=0.15,
        length_m=10.0,
        rise_m=0.0,
        fittings=[Fitting("elbow", 2)],
        equipment=[Equipment("EQ001", "AV", 24.0, rel_position=0.25)],
        waypoints=[(5.0, 0.0)],
    )
    active_nozzle = Nozzle("NZ001", "N002", "@/1", 0.00266666667)
    inactive_nozzle = Nozzle("NZ002", "N002", "@/2", 0.00266666667, status=0)
    valve = Valve("V001", "N001", "N002", "pressure-drop", target_value=0.0)

    network.add_node(node_1)
    network.add_node(node_2)
    network.add_pipe(pipe)
    network.add_nozzle(active_nozzle)
    network.add_nozzle(inactive_nozzle)
    network.add_valve(valve)

    assert network.get_node("N001") == node_1
    assert network.pipes["P001"] == pipe
    assert network.valves["V001"] == valve
    assert network.active_nozzles() == [active_nozzle]


def test_pipe_network_rejects_duplicate_node_and_pipe_ids() -> None:
    network = PipeNetwork(title="Duplicates")
    network.add_node(Node("N001", 0.0, 0.0, 0.0, "pipe"))
    network.add_node(Node("N002", 1.0, 0.0, 0.0, "pipe"))
    network.add_pipe(Pipe("P001", "N001", "N002", 0.15, 1.0, 0.0))

    with pytest.raises(ValueError, match="Duplicate node_id"):
        network.add_node(Node("N001", 2.0, 0.0, 0.0, "pipe"))

    with pytest.raises(ValueError, match="Duplicate pipe_id"):
        network.add_pipe(Pipe("P001", "N001", "N002", 0.15, 1.0, 0.0))


def test_pipe_network_exports_dataframes() -> None:
    network = PipeNetwork(title="Sample network")
    network.add_node(Node("N001", 0.0, 0.0, 48.55, "pipe", metadata={"floor": "4F"}))
    network.add_node(Node("N002", 10.0, 0.0, 48.55, "head"))
    network.add_pipe(
        Pipe(
            pipe_id="P001",
            from_node="N001",
            to_node="N002",
            diameter_m=0.15,
            length_m=10.0,
            rise_m=0.0,
            fittings=[Fitting("tee", 1)],
            metadata={"source": "plan"},
        )
    )
    network.add_nozzle(Nozzle("NZ001", "N002", "@/1", 0.00266666667))

    node_df = network.to_node_dataframe()
    pipe_df = network.to_pipe_dataframe()
    nozzle_df = network.to_nozzle_dataframe()
    fitting_df = network.to_fitting_dataframe()

    assert list(node_df["node_id"]) == ["N001", "N002"]
    assert node_df.loc[0, "floor"] == "4F"
    assert pipe_df.loc[0, "pipe_id"] == "P001"
    assert pipe_df.loc[0, "fitting_count"] == 1
    assert pipe_df.loc[0, "source"] == "plan"
    assert nozzle_df.loc[0, "library_item"] == "SP-HEAD"
    assert fitting_df.loc[0, "pipe_id"] == "P001"
    assert fitting_df.loc[0, "fitting_type"] == "tee"
